"""공문 개별발송: 연락처 이름이 거래처 검색에 포함되고 수신메일에 반영."""

import pandas as pd

from datetime import date

from price_increase_tab import (
    _apply_amount_to_items,
    _apply_pct_to_items,
    _items_df_for_editor,
    _increase_mode_is_pct,
    _compute_latest_unit_prices,
    default_increase_price,
    _BODY_CELL_ORDER,
    _DEFAULT_LETTER_PARAS,
    _body_text_to_paras,
    _clean_client_label,
    _norm_client_list,
    _person_key,
    _pi_attach_letter_pdf,
    _plain_mail_intro,
    _strip_letter_body_from_kwargs,
    compose_addrs_from_fields,
    cc_for_plain_mail,
    clients_matching_email,
    exclude_blocked_emails,
    keep_existing_mail_compose,
    filter_letter_client_names,
    hold_letter_preview_fields,
    join_recipient_emails,
    merge_keep_manual_emails,
    remove_email_addr,
    list_clients_for_letter,
    list_clients_for_staff,
    list_mail_contact_names,
    lookup_email_with_meta,
    lookup_emails_for_clients,
    resolve_sales_client_name,
)


def test_clean_client_label_strips_csv_quotes():
    assert _clean_client_label("'김온전'") == "김온전"
    assert _clean_client_label('"디아이지에어가스(주)"') == "디아이지에어가스(주)"
    assert _norm_client_list(["'김온전'", '"김온전"']) == ["김온전"]


def test_pct_increase_matches_base_times_rate():
    assert _increase_mode_is_pct("퍼센테이지(%)") is True
    assert _increase_mode_is_pct("인상금액(원)") is False
    assert default_increase_price(225.0, 10.0) == 247.5
    assert default_increase_price(22000.0, 10.0) == 24200.0
    rows = [
        {"선택": True, "품목명": "N2 (kg, Bulk)", "기존단가": 225.0, "인상적용단가": 225.0},
        {"선택": True, "품목명": "CO2", "기존단가": 22000.0, "인상적용단가": 22000.0},
        {"선택": False, "품목명": "skip", "기존단가": 100.0, "인상적용단가": 100.0},
    ]
    out = _apply_pct_to_items(rows, 10.0, only_selected=True)
    assert out[0]["인상적용단가"] == 247.5
    assert out[1]["인상적용단가"] == 24200.0
    assert out[2]["인상적용단가"] == 100.0
    amt = _apply_amount_to_items(rows, 10.0, only_selected=True)
    assert amt[0]["인상적용단가"] == 235.0


def test_person_key_strips_titles():
    assert _person_key("김도엽이사님") == _person_key("김도엽이사")
    assert _person_key("김도엽이사님") == "김도엽"
    assert _person_key("김영환 과장") == "김영환"


def test_letter_options_include_contact_names():
    sales = pd.DataFrame(
        {
            "담당자": ["김명현", "김명현"],
            "거래처": ["국제산업가스", "금정종합가스"],
        }
    )
    mail = pd.DataFrame(
        {
            "거래처": ["김도엽이사님", "김명현이사", "국제산업가스"],
            "이메일": ["doyeob@hanmail.net", "kimmh2580@gmail.com", "a@b.com"],
        }
    )
    staff_only = list_clients_for_staff(sales, "전체")
    assert "김도엽이사님" not in staff_only
    opts = list_clients_for_letter(sales, "전체", mail)
    assert "국제산업가스" in opts
    assert "김도엽이사님" in opts
    assert "김명현이사" in opts


def test_racool_alias_maps_to_sales_name_and_prices():
    """연락처 '라쿨'은 상단 검색과 같은 매출 정식명으로 맞추고 품목이 나와야 한다."""
    sales = pd.DataFrame(
        {
            "담당자": ["김명현", "김명현", "김명현"],
            "거래처": ["라쿨 주식회사(구.신정우)"] * 3,
            "품목명": ["N2", "O2", "Ar"],
            "단가": [100.0, 200.0, 300.0],
            "출고량": [10.0, 5.0, 1.0],
            "매출일": ["2026-11-01", "2026-11-02", "2026-11-03"],
        }
    )
    mail = pd.DataFrame(
        {
            "거래처": ["라쿨", "라쿨주식회사(구.신정우)", "김도엽이사님"],
            "이메일": ["jwko@racool.co.kr", "kskim@racool.co.kr", "a@b.com"],
        }
    )
    assert resolve_sales_client_name("라쿨", sales) == "라쿨 주식회사(구.신정우)"
    assert resolve_sales_client_name("라쿨주식회사(구.신정우)", sales) == "라쿨 주식회사(구.신정우)"
    opts = list_clients_for_letter(sales, "전체", mail)
    assert "라쿨 주식회사(구.신정우)" in opts
    assert "라쿨" not in opts
    assert "라쿨주식회사(구.신정우)" not in opts
    prices = _compute_latest_unit_prices(sales, "라쿨")
    assert set(prices["품목명"]) == {"N2", "O2", "Ar"}


def test_lookup_fills_email_for_title_variants():
    mail = pd.DataFrame(
        {
            "거래처": ["김도엽이사님", "김영환 과장"],
            "이메일": ["doyeob@hanmail.net", "01091910367@naver.com"],
        }
    )
    em, matched = lookup_email_with_meta("김도엽이사님", mail)
    assert em == "doyeob@hanmail.net"
    assert matched == "김도엽이사님"
    em2, matched2 = lookup_email_with_meta("김도엽이사", mail)
    assert em2 == "doyeob@hanmail.net"
    assert "김도엽" in matched2
    em3, _ = lookup_email_with_meta("김영환과장", mail)
    assert em3 == "01091910367@naver.com"
    quoted = pd.DataFrame({"거래처": ["'가람플랜트'"], "이메일": ["kim@g.com"]})
    em4, as4 = lookup_email_with_meta("가람플랜트", quoted)
    assert em4 == "kim@g.com"
    assert as4 == "가람플랜트"


def test_mail_contact_names_skips_empty():
    mail = pd.DataFrame({"거래처": ["김도엽이사님", "", "nan"], "이메일": ["a@b.com", "c@d.com", "e@f.com"]})
    assert list_mail_contact_names(mail) == ["김도엽이사님"]


def test_join_and_lookup_multiple_clients():
    assert join_recipient_emails(["a@b.com", "a@b.com", "c@d.com"]) == "a@b.com, c@d.com"
    assert _norm_client_list(["김도엽이사님", "김도엽이사님", " 국제산업가스 "]) == [
        "김도엽이사님",
        "국제산업가스",
    ]
    mail = pd.DataFrame(
        {
            "거래처": ["김도엽이사님", "국제산업가스", "메일없음"],
            "이메일": ["doyeob@hanmail.net", "a@b.com", ""],
        }
    )
    joined, matched, missing = lookup_emails_for_clients(
        ["김도엽이사", "국제산업가스", "없는곳"], mail
    )
    assert joined == "doyeob@hanmail.net, a@b.com"
    assert "김도엽이사님" in matched
    assert "없는곳" in missing


def test_strip_letter_body_from_attachment_kwargs():
    paras = _body_text_to_paras(_DEFAULT_LETTER_PARAS["C16"] + "\n다음줄")
    assert any(str(paras.get(c) or "").strip() for c in _BODY_CELL_ORDER)
    out = _strip_letter_body_from_kwargs(
        {"letter_paras": paras, "body": "메일본문", "items": [{"품목명": "CO2"}]}
    )
    assert out["body"] == ""
    assert out["items"] == [{"품목명": "CO2"}]
    assert all(str(out["letter_paras"].get(c) or "") == "" for c in _BODY_CELL_ORDER)
    assert str(out["letter_paras"].get("C35") or "").strip()


def test_filter_letter_client_names_keeps_list_small():
    names = ["국제산업가스", "금정종합가스", "김도엽이사님", "김명현이사", "김영환 과장"]
    hits = filter_letter_client_names(names, "김도엽이사", exclude=[], limit=12)
    assert hits == ["김도엽이사님"]
    hits2 = filter_letter_client_names(names, "김", exclude=["김도엽이사님"], limit=12)
    assert "김도엽이사님" not in hits2
    assert "김명현이사" in hits2
    assert filter_letter_client_names(names, "", exclude=[], limit=12) == []
    similar = filter_letter_client_names(
        ["가람플랜트", "국제산업가스", "금정종합가스"], "가람", exclude=[], limit=12
    )
    assert similar[0] == "가람플랜트"


def test_plain_mail_skips_letter_pdf():
    assert _pi_attach_letter_pdf(True) is True
    assert _pi_attach_letter_pdf(False) is False
    assert _plain_mail_intro("김도엽이사님") == "김도엽이사님 귀중\n\n"
    assert _plain_mail_intro("") == ""


def test_exclude_blocked_auto_email_keeps_manual():
    assert exclude_blocked_emails("a@x.com, b@y.com", ["a@x.com"]) == "b@y.com"
    assert exclude_blocked_emails("a@x.com", ["A@x.com"]) == ""
    assert exclude_blocked_emails("a@x.com, b@y.com", []) == "a@x.com, b@y.com"
    merged = merge_keep_manual_emails("a@x.com", "b@y.com", "b@y.com")
    assert exclude_blocked_emails(merged, ["b@y.com"]) == "a@x.com"


def test_remove_email_updates_compose_join():
    left = remove_email_addr("3023526@gmail.com, ecount@ecounterp.com", "ecount@ecounterp.com")
    assert left == "3023526@gmail.com"
    to_addr, cc_addr = compose_addrs_from_fields(left, "3023526@gmail.com")
    assert to_addr == "3023526@gmail.com"
    assert cc_addr == "3023526@gmail.com"


def test_compose_follows_typed_to_and_cc():
    to_addr, cc_addr = compose_addrs_from_fields(
        "3023526@gmail.com, 3023526@gmail.com",
        "cc@x.com, other@y.com",
    )
    assert to_addr == "3023526@gmail.com"
    assert cc_addr == "cc@x.com, other@y.com"
    to2, cc2 = compose_addrs_from_fields("3033526@gmail.com", "")
    assert to2 == "3033526@gmail.com"
    assert cc2 == ""


def test_cc_only_for_plain_mail():
    assert cc_for_plain_mail("a@b.com", attach_letter=True) == ""
    assert cc_for_plain_mail("a@b.com", attach_letter=False) == "a@b.com"
    assert cc_for_plain_mail("  ", attach_letter=False) == ""


def test_remove_email_addr_and_matching_clients():
    assert remove_email_addr("a@x.com, b@y.com", "a@x.com") == "b@y.com"
    assert remove_email_addr("a@x.com", "A@x.com") == ""
    assert remove_email_addr("a@x.com, b@y.com", "") == "a@x.com, b@y.com"
    mail = pd.DataFrame(
        {"거래처": ["가람플랜트", "남효윤"], "이메일": ["a@x.com", "b@y.com"]}
    )
    assert clients_matching_email(["가람플랜트", "남효윤"], mail, "a@x.com") == [
        "가람플랜트"
    ]
    assert clients_matching_email(["가람플랜트"], mail, "other@z.com") == []


def test_merge_manual_and_client_emails_both_ways():
    assert merge_keep_manual_emails("a@x.com", "", "b@y.com") == "a@x.com, b@y.com"
    assert merge_keep_manual_emails("b@y.com", "b@y.com", "b@y.com, c@z.com") == "b@y.com, c@z.com"
    assert merge_keep_manual_emails("b@y.com, a@x.com", "b@y.com", "b@y.com") == "b@y.com, a@x.com"
    assert merge_keep_manual_emails("b@y.com, a@x.com", "b@y.com", "") == "a@x.com"
    assert merge_keep_manual_emails("b@y.com, a@x.com", "b@y.com", "c@z.com") == "a@x.com, c@z.com"
    assert merge_keep_manual_emails("a@x.com, a@x.com", "", "a@x.com") == "a@x.com"


def test_typed_addr_source_keeps_enter_value_after_form_clear():
    from unittest import mock

    import price_increase_tab as m

    ss = {"pi_to_add": "", "pi_cc_add": "", "pi_to_add_commit": "", "pi_cc_add_commit": ""}
    with mock.patch.object(m.st, "session_state", ss):
        assert m._typed_addr_source("pi_to_add", "recv@mail.com") == "recv@mail.com"
        assert m._typed_addr_source("pi_cc_add", "cc@mail.com") == "cc@mail.com"
    ss["pi_cc_add"] = ""
    ss["pi_cc_add_commit"] = "cc@mail.com"
    with mock.patch.object(m.st, "session_state", ss):
        assert m._typed_addr_source("pi_cc_add", "") == "cc@mail.com"


def test_take_typed_addrs_updates_to_and_cc_separately():
    from unittest import mock

    import price_increase_tab as m

    ss = {
        "pi_single_email": "to@x.com",
        "pi_single_cc": "",
        "pi_mail_draft": {"to": "to@x.com", "cc": ""},
    }
    with mock.patch.object(m.st, "session_state", ss):
        assert m._take_typed_addrs("cc@y.com", "pi_single_cc") is True
        assert ss["pi_single_cc"] == "cc@y.com"
        assert ss["pi_mail_draft"]["cc"] == "cc@y.com"
        assert ss["pi_mail_draft"]["to"] == "to@x.com"
        assert ss["pi_mail_compose_cc_view"] == "cc@y.com"
        assert m._take_typed_addrs("to2@x.com", "pi_single_email") is True
        assert ss["pi_single_email"] == "to@x.com, to2@x.com"
        assert ss["pi_mail_draft"]["cc"] == "cc@y.com"


def test_mail_ipad_css_portrait_stacks_landscape_keeps_row():
    from pathlib import Path

    src = Path(__file__).resolve().parent.joinpath("price_increase_tab.py").read_text(
        encoding="utf-8"
    )
    assert "@media (max-width:1180px) and (orientation:portrait)" in src
    assert "@media (min-width:851px) and (max-width:1180px) and (orientation:landscape)" in src
    assert "st-key-pi_mail_inline_send" in src


def test_finish_mail_send_stays_on_compose():
    from unittest import mock

    import price_increase_tab as m

    ss = {
        "pi_left_mode": "mail",
        "pi_left_opened": True,
        "pi_mail_draft": {"to": "a@b.com"},
    }
    with mock.patch.object(m.st, "session_state", ss):
        m._finish_pi_mail_send(ok=True, msg="발송 완료", to_addr="a@b.com")
        assert ss["pi_left_mode"] == "mail"
        assert ss["pi_keep_mail_after_send"] is True
        assert ss["pi_send_flash"]["ok"] is True
        assert m._pi_consume_left_nav_block() is True
        ss["pi_left_mode"] = "summary"
        m._ensure_mail_compose_first_screen({})
        assert ss["pi_left_mode"] == "mail"
        assert m._pi_consume_left_nav_block() is False


def test_seed_compose_widgets_restores_after_preview():
    from unittest import mock

    import price_increase_tab as m

    ss = {
        "pi_mail_draft": {
            "subject": "[신일가스] 가스가격 조정의 건",
            "body": "미리보기 전에 쓴 본문",
            "to": "a@b.com",
        },
        "pi_mail_compose_subject": "",
        "pi_mail_compose_body": "",
        "pi_mail_text_edited": True,
    }
    with mock.patch.object(m.st, "session_state", ss):
        assert m._mail_compose_in_progress() is True
        m._pi_seed_compose_widgets_from_draft()
    assert ss["pi_mail_compose_subject"] == "[신일가스] 가스가격 조정의 건"
    assert ss["pi_mail_compose_body"] == "미리보기 전에 쓴 본문"
    src = open("price_increase_tab.py", encoding="utf-8").read()
    left = src.split("def _render_pi_left_panel", 1)[1].split("\ndef ", 1)[0]
    assert "if do_preview:" in left
    assert left.split("if do_preview:", 1)[1].split("if do_send:", 1)[0].count("_remember_mail_compose_text()") >= 1
    assert left.split("if do_send:", 1)[1].split("if do_summary:", 1)[0].count("_remember_mail_compose_text()") >= 1


def test_remember_mail_compose_text_keeps_subject_and_body():
    import os
    import tempfile
    from unittest import mock

    import price_increase_tab as m

    path = os.path.join(tempfile.mkdtemp(), "compose_live.json")
    ss = {
        "pi_mail_draft": {"subject": "옛 제목", "body": "옛 본문", "to": "a@b.com"},
        "pi_mail_compose_subject": "방문 안내",
        "pi_mail_compose_body": "내일 오전 미팅입니다.",
    }
    with mock.patch.object(m, "PI_COMPOSE_LIVE", path), mock.patch.object(m.st, "session_state", ss):
        m._remember_mail_compose_text()
        assert ss["pi_mail_draft"]["subject"] == "방문 안내"
        assert ss["pi_mail_draft"]["body"] == "내일 오전 미팅입니다."
        assert ss["pi_mail_draft"]["to"] == "a@b.com"
        assert ss["pi_mail_text_edited"] is True
        assert m._mail_compose_in_progress() is True
        ss.pop("pi_mail_compose_subject")
        ss.pop("pi_mail_compose_body")
        assert m._mail_compose_in_progress() is True


def test_compose_live_survives_reload():
    import os
    import tempfile
    from unittest import mock

    import price_increase_tab as m

    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "compose_live.json")
    m.save_compose_live(
        {"subject": "제품가격 조정의 건", "body": "내일 방문합니다.", "to": "a@b.com", "cc": "", "client": "경민"},
        path,
    )
    loaded = m.load_compose_live(path)
    assert loaded["subject"] == "제품가격 조정의 건"
    assert loaded["body"] == "내일 방문합니다."
    ss = {}
    with mock.patch.object(m, "PI_COMPOSE_LIVE", path), mock.patch.object(m.st, "session_state", ss):
        m._restore_compose_live()
        assert ss["pi_mail_compose_subject"] == "제품가격 조정의 건"
        assert ss["pi_mail_compose_body"] == "내일 방문합니다."
        assert ss["pi_mail_draft"]["subject"] == "제품가격 조정의 건"
        m._restore_compose_live()
        ss["pi_mail_compose_subject"] = "수정한 제목"
        m._restore_compose_live()
        assert ss["pi_mail_compose_subject"] == "수정한 제목"
    m.clear_compose_live(path)
    assert m.load_compose_live(path) == {}


def test_hold_edit_opens_in_mail_compose():
    import os
    import tempfile
    from unittest import mock

    import price_increase_tab as m

    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "compose_live.json")
    ss = {
        "pi_hold_edit_pending": {
            "id": "hold1",
            "subject": "[신일가스] 단가조정의 건",
            "body": "고칠 본문",
            "to": "nksteel@hanmail.net",
            "cc": "",
            "client": "신일가스",
            "attach_letter": True,
            "letter_kwargs": {
                "client": "신일가스",
                "title": "제품가격 조정의 건",
                "items": [{"선택": True, "품목명": "N2", "기존단가": 100, "인상적용단가": 120}],
                "letter_paras": {
                    "C16": "1. 감사",
                    "C18": "2. 안내",
                    "C19": "3. 조정",
                    "C21": "4. 부탁",
                },
            },
        }
    }
    with mock.patch.object(m, "PI_COMPOSE_LIVE", path), mock.patch.object(m.st, "session_state", ss):
        m._consume_hold_edit_open()
    assert ss["pi_mail_compose_subject"] == "[신일가스] 단가조정의 건"
    assert ss["pi_mail_compose_body"] == "고칠 본문"
    assert ss["pi_hold_editing_id"] == "hold1"
    assert ss["pi_mail_draft"]["hold_id"] == "hold1"
    assert ss["pi_single_email"] == "nksteel@hanmail.net"
    assert ss["pi_single_clients"] == ["신일가스"]
    assert ss["pi_hold_open_mail_tab"] is True
    assert "pi_fresh_mail" not in ss
    restore = ss["pi_hold_letter_restore"]
    assert restore["items"][0]["인상적용단가"] == 120
    assert "3. 조정" in restore["letter_body"]
    with mock.patch.object(m.st, "session_state", ss):
        m._apply_hold_letter_restore(
            client="신일가스",
            body_key="pi_letter_body_신일가스",
            items_key="pi_items_신일가스",
        )
    assert ss["pi_items_신일가스"][0]["인상적용단가"] == 120
    assert "3. 조정" in ss["pi_letter_body_신일가스"]
    assert ss["pi_single_title"] == "제품가격 조정의 건"
    assert "def _pi_click_inner_tab" in open("price_increase_tab.py", encoding="utf-8").read()
    assert ss["pi_hold_open_mail_tab"] is True
    assert "pi_hold_edit_pending" not in ss


def test_update_hold_mail_replaces_same_row():
    import os
    import tempfile

    from price_increase_tab import append_hold_mail, load_hold_mail, update_hold_mail

    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "hold_mail.json")
    saved = append_hold_mail(
        {
            "to": "nksteel@hanmail.net",
            "cc": "",
            "subject": "[신일가스] 단가조정의 건",
            "body": "이전 본문",
            "client": "신일가스",
            "letter_kwargs": {"title": "단가조정의 건"},
            "send": True,
        },
        path,
    )
    other = append_hold_mail(
        {"to": "djsy7615@naver.com", "subject": "다른 메일", "body": "유지", "send": True},
        path,
    )
    updated = update_hold_mail(
        saved["id"],
        {"subject": "[신일가스] 가격인상 조정의 건", "body": "고친 본문", "to": "nksteel@hanmail.net", "cc": "a@b.com"},
        path,
    )
    rows = load_hold_mail(path)
    assert updated is not None
    assert updated["id"] == saved["id"]
    assert len(rows) == 2
    assert rows[0]["id"] == saved["id"]
    assert rows[0]["subject"] == "[신일가스] 가격인상 조정의 건"
    assert rows[0]["body"] == "고친 본문"
    assert rows[0]["cc"] == "a@b.com"
    assert rows[0]["letter_kwargs"]["title"] == "단가조정의 건"
    assert rows[0]["ts"] != saved["ts"]
    assert rows[1]["id"] == other["id"]
    assert rows[1]["body"] == "유지"
    assert update_hold_mail("없는id", {"subject": "x"}, path) is None


def test_rehydrate_keeps_price_increase_subject_after_restart():
    import os
    import tempfile
    from unittest import mock

    import price_increase_tab as m

    path = os.path.join(tempfile.mkdtemp(), "compose_live.json")
    m.save_compose_live(
        {
            "subject": "[신일가스] 가격인상 조정의 건",
            "body": "내일 방문합니다.",
            "to": "dreamlaser@naver.com",
            "cc": "",
        },
        path,
    )
    ss = {}
    with mock.patch.object(m, "PI_COMPOSE_LIVE", path), mock.patch.object(m.st, "session_state", ss):
        m._rehydrate_compose_text()
    assert ss["pi_mail_compose_subject"] == "[신일가스] 가격인상 조정의 건"
    assert ss["pi_mail_compose_body"] == "내일 방문합니다."
    assert ss["pi_mail_draft"]["subject"] == "[신일가스] 가격인상 조정의 건"
    ss.pop("pi_mail_compose_subject")
    ss.pop("pi_mail_compose_body")
    with mock.patch.object(m, "PI_COMPOSE_LIVE", path), mock.patch.object(m.st, "session_state", ss):
        m._open_mail_compose_dialog(
            client="드림레이저",
            email="",
            title="제품가격 조정의 건",
            body="",
            pdf_name="공문.pdf",
            letter_kwargs={},
            keep_compose=False,
        )
    assert ss["pi_mail_draft"]["subject"] == "[신일가스] 가격인상 조정의 건"
    assert ss["pi_mail_draft"]["body"] == "내일 방문합니다."


def test_fresh_mail_drops_previous_recipients():
    from unittest import mock

    import price_increase_tab as m

    ss = {
        "pi_fresh_mail": True,
        "pi_single_email": "dreamlaser@naver.com, mskt2020@naver.com",
        "pi_single_cc": "old@b.com",
        "pi_email_auto": "dreamlaser@naver.com",
        "pi_single_clients": ["드림레이저"],
        "pi_single_client": "드림레이저",
        "pi_hold_editing_id": "hold1",
        "pi_skip_email_fill": True,
        "pi_mail_draft": {"to": "dreamlaser@naver.com", "cc": "old@b.com", "subject": "제목", "body": "본문"},
    }
    with mock.patch.object(m.st, "session_state", ss), mock.patch.object(m, "save_compose_live"):
        m._pi_apply_fresh_mail()
    assert ss["pi_single_email"] == ""
    assert ss["pi_single_cc"] == ""
    assert ss["pi_single_clients"] == []
    assert ss["pi_mail_draft"]["to"] == ""
    assert ss["pi_mail_draft"]["subject"] == "제목"
    assert "pi_fresh_mail" not in ss
    assert "pi_hold_editing_id" not in ss
    src = open("price_increase_tab.py", encoding="utf-8").read()
    hold = src.split("elif do_hold", 1)[1].split("elif do_cancel", 1)[0]
    assert 'st.session_state["pi_fresh_mail"] = True' in hold


def test_new_client_selection_drops_previous_recipients():
    from unittest import mock

    import price_increase_tab as m

    ss = {
        "pi_email_client": "드림레이저",
        "pi_single_email": "dreamlaser@naver.com, mskt2020@naver.com",
        "pi_email_auto": "",
        "pi_mail_draft": {"to": "dreamlaser@naver.com, mskt2020@naver.com"},
        "pi_single_clients": ["엔케이철강"],
        "_pi_client_alias": {},
        "_pi_mail_df_cache": None,
    }
    with mock.patch.object(m.st, "session_state", ss), mock.patch.object(
        m, "lookup_emails_for_clients", return_value=("nksteel@hanmail.net", ["엔케이철강"], [])
    ):
        m._pi_on_clients_change()
    assert ss["pi_single_email"] == "nksteel@hanmail.net"
    assert "dreamlaser@naver.com" not in ss["pi_single_email"]
    assert "mskt2020@naver.com" not in ss["pi_single_email"]


def test_hold_letter_preview_fields_from_kwargs():
    assert hold_letter_preview_fields({"body": "메일만", "attach_letter": False}) is None
    prev = hold_letter_preview_fields(
        {
            "client": "태광기업",
            "attach_letter": True,
            "letter_kwargs": {
                "client": "태광기업",
                "title": "제품가격 조정의 건",
                "doc_no": "신일(1)-2611-0000",
                "send_date": "2026-11-26",
                "effective": "2026-12-01",
                "letter_paras": {"C16": "1. 귀사의 일익 번창하심을 기원하며"},
                "items": [{"품목명": "N2", "기존단가": 100, "인상적용단가": 110}],
            },
        }
    )
    assert prev is not None
    assert prev["client"] == "태광기업"
    assert prev["title"] == "제품가격 조정의 건"
    assert prev["doc_no"] == "신일(1)-2611-0000"
    assert prev["send_date"] == date(2026, 11, 26)
    assert "1. 귀사의 일익 번창하심을 기원하며" in prev["letter_body"]
    assert prev["items"][0]["품목명"] == "N2"


def test_hold_save_refreshes_hold_list():
    src = open("price_increase_tab.py", encoding="utf-8").read()
    hold = src.split("elif do_hold", 1)[1].split("elif do_cancel", 1)[0]
    assert "pi_hold_flash" in hold
    assert "pi_hold_show" in hold
    assert "_pi_rerun(full=True)" in hold
    assert "def _pi_focus_hold_tab" in src
    hold_fn = src.split("def _render_pi_hold_mail", 1)[1].split("\ndef ", 1)[0]
    assert "pi_hold_open_" in hold_fn
    assert "on_click=_pi_hold_open_subject" in hold_fn
    assert "_render_pi_live_letter_page" in hold_fn
    assert "hold_letter_preview_fields" in hold_fn
    assert "메일 내용" not in hold_fn
    assert "pi_hold_body_" not in hold_fn
    assert "st.selectbox" not in hold_fn
    assert "메일을 선택하면" not in hold_fn
    assert "compact=True" in hold_fn
    assert "제목을 누르면 첨부 공문이 열립니다" in hold_fn
    assert 'st.button("수정"' in hold_fn
    assert "pi_hold_edit_{hold_id}" in hold_fn
    assert "저장시각" in hold_fn
    assert "hold_mail_label(row)" not in hold_fn.split("if not rows", 1)[-1].split("by_id", 1)[0]
    assert "pi_hold_edit_pending" in hold_fn
    assert "save_compose_hold" in hold
    compose = src.split("def _render_pi_left_mail_compose", 1)[1].split("\ndef ", 1)[0]
    assert '"다시 임시저장" if editing_hold else "임시저장"' in compose


def test_first_hold_appends_and_edit_updates_same_mail():
    import os
    import tempfile

    from price_increase_tab import load_hold_mail, save_compose_hold

    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "hold_mail.json")
    first, updated = save_compose_hold(
        subject="처음 제목",
        body="처음 본문",
        to="a@b.com",
        client="경민산업",
        path=path,
    )
    assert updated is False
    rows = load_hold_mail(path)
    assert len(rows) == 1
    assert rows[0]["id"] == first["id"]
    assert rows[0]["subject"] == "처음 제목"
    again, updated = save_compose_hold(
        edit_id=first["id"],
        subject="고친 제목",
        body="고친 본문",
        to="a@b.com",
        client="경민산업",
        path=path,
    )
    rows = load_hold_mail(path)
    assert updated is True
    assert len(rows) == 1
    assert again["id"] == first["id"]
    assert rows[0]["subject"] == "고친 제목"
    assert rows[0]["body"] == "고친 본문"


def test_bind_manual_email_to_client_without_contact():
    import os
    import tempfile
    from unittest import mock

    import pandas as pd

    import price_increase_tab as m

    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "mail_contacts.csv")
    empty = pd.DataFrame(columns=["거래처", "이메일", "비고"])
    with mock.patch.object(m, "PI_MAIL_CSV", path), mock.patch.object(
        m, "load_mail_contacts", return_value=empty
    ):
        df, saved = m.bind_manual_emails_to_clients(
            "blt@example.com",
            ["비엘티(구.에스앤케이)"],
            mail_df=empty,
            note="테스트",
        )
    assert saved == ["비엘티(구.에스앤케이)"]
    assert os.path.isfile(path)
    loaded = pd.read_csv(path)
    assert list(loaded["거래처"]) == ["비엘티(구.에스앤케이)"]
    assert list(loaded["이메일"]) == ["blt@example.com"]
    # 이미 있으면 다시 안 덮음
    df2, saved2 = m.bind_manual_emails_to_clients(
        "other@example.com",
        ["비엘티(구.에스앤케이)"],
        mail_df=df,
    )
    assert saved2 == []
    assert list(df2["이메일"]) == ["blt@example.com"]


def test_hold_client_prefers_selected_over_stale_draft_name():
    import os
    import tempfile

    from price_increase_tab import (
        delete_hold_mail,
        load_hold_mail,
        resolve_hold_client,
        save_compose_hold,
    )

    assert (
        resolve_hold_client(
            client="일광잉크페인트",
            letter_kwargs={"client": "비엘티(구.에스앤케이)"},
            to="3023526@gmail.com",
            picked=["비엘티(구.에스앤케이)"],
        )
        == "비엘티(구.에스앤케이)"
    )
    assert (
        resolve_hold_client(
            client="일광잉크페인트",
            letter_kwargs={"client": "비엘티(구.에스앤케이)"},
            to="3023526@gmail.com",
            picked=[],
        )
        == "비엘티(구.에스앤케이)"
    )
    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "hold_mail.json")
    saved, updated = save_compose_hold(
        subject="가격인상",
        body="본문",
        to="3023526@gmail.com",
        client="일광잉크페인트",
        letter_kwargs={"client": "비엘티(구.에스앤케이)"},
        picked=["비엘티(구.에스앤케이)"],
        path=path,
    )
    assert updated is False
    assert saved["client"] == "비엘티(구.에스앤케이)"
    assert delete_hold_mail(saved["id"], path) is True
    assert load_hold_mail(path) == []
    src = open("price_increase_tab.py", encoding="utf-8").read()
    cancel = src.split("elif do_cancel", 1)[1].split("_render_pi_send_flash", 1)[0]
    assert "delete_hold_mail" in cancel


def test_mail_subject_body_not_wiped_on_tab_restore():
    src = open("app.py", encoding="utf-8").read()
    fn = src.split("def _dash_is_pi_widget_key", 1)[1].split("def ", 1)[0]
    assert "pi_mail_compose_subject" in fn
    assert "pi_mail_compose_body" in fn
    assert "return False" in fn


def test_keep_mail_compose_when_letter_reattached():
    title, body = keep_existing_mail_compose(
        title="단가인상 안내",
        body="기본 공문 메일",
        existing_subject="회의 안내",
        existing_body="내일 방문합니다.",
        keep=True,
    )
    assert title == "회의 안내"
    assert body == "내일 방문합니다."
    title2, body2 = keep_existing_mail_compose(
        title="단가인상 안내",
        body="기본 공문 메일",
        existing_subject="회의 안내",
        existing_body="내일 방문합니다.",
        keep=False,
    )
    assert title2 == "단가인상 안내"
    assert body2 == "기본 공문 메일"
    title3, body3 = keep_existing_mail_compose(
        title="단가인상 안내",
        body="기본",
        existing_subject="",
        existing_body="",
        keep=True,
    )
    assert title3 == ""
    assert body3 == ""


def test_editor_old_new_price_no_decimals():
    df = _items_df_for_editor(
        [{"품목명": "N2 (kg, Bulk)", "기존단가": 170.0, "인상적용단가": 200.0}]
    )
    assert int(df.loc[0, "기존단가"]) == 170
    assert int(df.loc[0, "인상단가"]) == 200
    src = open("price_increase_tab.py", encoding="utf-8").read()
    assert src.count('format="%d"') >= 2


def test_sent_log_shows_latest_client_first():
    import tempfile
    from pathlib import Path
    from unittest.mock import patch

    import price_increase_tab as pit

    with tempfile.TemporaryDirectory() as td:
        log = Path(td) / "sent_log.jsonl"
        log.write_text(
            '{"ts":"2026-09-11T12:00:00","date":"2026-09-11","client":"Z가람플랜트","email":"a@b.com","subject":"old","ok":true,"mode":"single","staff":"","items":1,"msg":"ok"}\n'
            '{"ts":"2026-10-29T09:11:04","date":"2026-10-29","client":"예성이엔지(아산)","email":"yesung@yesung-eng.com","subject":"가스 공급 단가 인상 안내의 건","ok":true,"mode":"single","staff":"","items":10,"msg":"ok"}\n',
            encoding="utf-8",
        )
        with patch.object(pit, "PI_SENT_LOG", str(log)):
            df = pit.load_sent_log()
        assert str(df.iloc[0]["client"]) == "예성이엔지(아산)"
        summary = pit.sent_summary_by_client(df)
        assert str(summary.iloc[0]["거래처"]) == "예성이엔지(아산)"
        assert pit._sent_log_body_text(df.iloc[0].get("body")) == ""


def test_client_search_finds_laser_spelling():
    from price_increase_tab import (
        _client_options_with_aliases,
        _client_spell_aliases,
        _match_client_names,
    )

    names = ["드림레이저", "대진에스에이", "엔케이철강"]
    assert _match_client_names(names, "드림레이져") == ["드림레이저"]
    assert _match_client_names(names, "드림레이저") == ["드림레이저"]
    assert "드림레이져" in _client_spell_aliases("드림레이저")
    opts, alias_of = _client_options_with_aliases(names)
    assert "드림레이저" in opts
    assert "드림레이져" in opts
    assert alias_of["드림레이져"] == "드림레이저"
    src = open("price_increase_tab.py", encoding="utf-8").read()
    picker = src.split("def _render_single_client_picker", 1)[1].split("\ndef ", 1)[0]
    assert 'placeholder="거래처명 입력"' in picker
    assert "st.multiselect(" in picker


def test_sent_log_stores_body_and_client_search():
    import tempfile
    from pathlib import Path
    from unittest.mock import patch

    import price_increase_tab as pit

    with tempfile.TemporaryDirectory() as td:
        log = Path(td) / "sent_log.jsonl"
        log.write_text(
            '{"ts":"2026-09-11T12:00:00","date":"2026-09-11","client":"Z가람플랜트","email":"a@b.com","subject":"old","ok":true,"mode":"single","staff":"","items":1,"msg":"ok"}\n',
            encoding="utf-8",
        )
        with (
            patch.object(pit, "PI_SENT_LOG", str(log)),
            patch.object(pit, "PI_DRAFTS", td),
            patch.object(pit, "PI_DIR", td),
        ):
            pit.append_sent_log(
                client="예성이엔지(아산)",
                email="yesung@yesung-eng.com",
                subject="가스 공급 단가 인상 안내의 건",
                ok=True,
                body="안녕하세요\n단가를 안내합니다.",
            )
            df = pit.load_sent_log()
        new = df[df["client"].astype(str) == "예성이엔지(아산)"].iloc[0]
        old = df[df["client"].astype(str) == "Z가람플랜트"].iloc[0]
        assert "단가를 안내합니다" in pit._sent_log_body_text(new["body"])
        assert pit._sent_log_body_text(old["body"]) == ""
        hit = pit.filter_sent_log_by_client(df, "(아산)")
        assert list(hit["client"].astype(str)) == ["예성이엔지(아산)"]
        assert pit.filter_sent_log_by_client(df, "없는거래처").empty
        show = pit._sent_history_show(hit)
        assert "body" not in show.columns
        assert list(show.columns)[:4] == ["발송일", "거래처", "제목", "이메일"]


def test_picked_letter_body_stays_when_client_changes():
    from price_increase_tab import _pi_keep_letter_body

    chosen = "1. 질소가스 공급부족으로 가격을 조정합니다."
    bodies = {
        "pi_letter_body_갑": chosen,
        "pi_letter_body_을": "1.귀사의 무궁한 발전을 기원하며",
    }
    kept = _pi_keep_letter_body(
        client="을",
        body_key="pi_letter_body_을",
        prev_client="갑",
        pinned=True,
        bodies=bodies,
    )
    assert kept == chosen
    fresh = _pi_keep_letter_body(
        client="을",
        body_key="pi_letter_body_을",
        prev_client="갑",
        pinned=False,
        bodies=bodies,
    )
    assert fresh == "1.귀사의 무궁한 발전을 기원하며"


def test_letter_spacing_fills_the_page_gap():
    from price_increase_tab import _PI_LINE_H, _pi_letter_block_spacing

    line_h, section_gap, table_gap, fits = _pi_letter_block_spacing(
        n_lines=10,
        n_section_gaps=3,
        n_tight_gaps=2,
        table_h=40.0,
        room=150.0,
        has_table=True,
    )
    assert fits
    assert abs(line_h - _PI_LINE_H) < 0.01
    assert table_gap < section_gap
    used = 10 * line_h + 2 * 0.6 + 3 * section_gap + table_gap + 40.0
    assert used < 150.0
    assert used > 120.0
    tight = _pi_letter_block_spacing(
        n_lines=40,
        n_section_gaps=8,
        n_tight_gaps=0,
        table_h=40.0,
        room=80.0,
        has_table=True,
    )
    assert tight[3] is False


def test_flow_letter_body_removes_blank_gaps():
    from price_increase_tab import _pi_flow_letter_body

    raw = (
        "1. 귀사의 일익 번창하심을 기원합니다.\n"
        "\n"
        "2. 당사는 안정적인 가스공급에 최선을\n"
        "다해왔으나 수급이 불안정합니다.\n"
        "3. 이에 공급가격을 조정합니다."
    )
    out = _pi_flow_letter_body(raw)
    assert "\n\n" not in out
    assert out.splitlines() == [
        "1. 귀사의 일익 번창하심을 기원합니다.",
        "2. 당사는 안정적인 가스공급에 최선을 다해왔으나 수급이 불안정합니다.",
        "3. 이에 공급가격을 조정합니다.",
    ]


def test_default_letter_body_matches_price_adjustment_form():
    from price_increase_tab import _default_letter_body_text, _default_letter_title

    body = _default_letter_body_text()
    assert _default_letter_title() == "제품가격 조정의 건"
    assert body.startswith("1. 귀사의 일익 번창하심을 기원하며")
    assert "\n\n2. 당사는 안정적인 가스공급에 최선을 다해왔으나" in body
    assert "\n\n4. 다시 한번 귀사의 일익 번창하심을 기원하며" in body
    assert "무궁한 발전" not in body


def test_take_letter_file_body_keeps_manual_edits():
    from price_increase_tab import _pi_take_letter_file_body

    assert _pi_take_letter_file_body(
        picked="공문261119 · 제품가격 조정의 건",
        applied="",
        token="b",
        applied_token="",
        current="옛 문장",
        previous_file_body="",
        new_body="1. 새 문장",
    )
    assert _pi_take_letter_file_body(
        picked="공문261119 · 제품가격 조정의 건",
        applied="공문261119 · 제품가격 조정의 건",
        token="new",
        applied_token="old",
        current="1. 파일 문장",
        previous_file_body="1. 파일 문장",
        new_body="1. 고친 파일 문장",
    )
    assert not _pi_take_letter_file_body(
        picked="공문261119 · 제품가격 조정의 건",
        applied="공문261119 · 제품가격 조정의 건",
        token="new",
        applied_token="old",
        current="1. 내가 고친 문장",
        previous_file_body="1. 파일 문장",
        new_body="1. 고친 파일 문장",
    )


def test_price_letter_book_title_fills_body():
    import os
    import tempfile
    from openpyxl import Workbook

    from price_increase_tab import _pi_letter_title_from_line, _pi_read_price_letters

    assert _pi_letter_title_from_line("제    목 : 제품가격 조정의 건") == "제품가격 조정의 건"
    wb = Workbook()
    ws = wb.active
    ws.title = "공문261119"
    ws["C12"] = "제    목 : 제품가격 조정의 건"
    ws["C16"] = "1. 귀사의 일익 번창하심을 기원합니다."
    ws["C18"] = "2. 질소가스 공급부족으로 가격을 조정합니다."
    ws["C29"] = "- 아     래 -"
    ws["C31"] = "품명"
    fd, path = tempfile.mkstemp(suffix=".xlsx")
    os.close(fd)
    try:
        wb.save(path)
        rows = _pi_read_price_letters(path)
    finally:
        os.remove(path)
    assert len(rows) == 1
    assert rows[0]["title"] == "제품가격 조정의 건"
    assert rows[0]["label"].endswith("· 제품가격 조정의 건")
    assert "질소가스" in rows[0]["body"]
    assert "품명" not in rows[0]["body"]
    assert "아래" not in rows[0]["body"]


def test_letter_folder_lists_every_xlsx():
    import os
    import tempfile

    from price_increase_tab import _pi_letter_book_paths_in, _pi_read_price_letter_files

    folder = tempfile.mkdtemp()
    names = [
        "수소솟테이지 공문.xlsx",
        "채권독촉공문.xlsx",
        "탄산단가인상공문.xlsx",
        "26년 가격조정공문_20261119.xlsx",
        "26년 가격조정공문_20261119(최종).xlsx",
        "~$채권독촉공문.xlsx",
    ]
    try:
        from openpyxl import Workbook

        for i, name in enumerate(names):
            wb = Workbook()
            ws = wb.active
            ws.title = "본문"
            ws["C12"] = f"제    목 : 안내 {i}"
            ws["C16"] = "1. 본문입니다."
            if "채권" in name:
                ws["C18"] = "- 아 래 -"
                ws["C20"] = "미수 총액 : 금 1원"
            path = os.path.join(folder, name)
            wb.save(path)
            os.utime(path, (1_700_000_000 + i, 1_700_000_000 + i))
        paths = _pi_letter_book_paths_in(folder)
        assert [os.path.basename(p) for p in paths] == list(reversed(names[:-1]))
        rows = _pi_read_price_letter_files(paths)
        labels = [row["label"] for row in rows]
        assert len(labels) == 5
        assert any("수소솟테이지 공문 · 안내" in label for label in labels)
        assert any("채권독촉공문 · 안내" in label for label in labels)
        assert any("가격조정공문_20261119(최종) · 안내" in label for label in labels)
        bond = next(row for row in rows if row["label"].startswith("채권독촉공문"))
        assert "미수 총액" in bond["body"]
        older = os.path.join(folder, "26년 가격조정공문_20261119.xlsx")
        newer = os.path.join(folder, "26년 가격조정공문_20261119(최종).xlsx")
        from openpyxl import load_workbook as _load

        for path, line in ((older, "1. 예전 문장"), (newer, "1. 최종 문장")):
            book = _load(path)
            book.active["C16"] = line
            book.save(path)
            book.close()
        merged = _pi_read_price_letter_files([newer, older])
        same = [row for row in merged if row.get("key") == "본문|안내 4"]
        assert len(same) == 1
        assert "최종 문장" in same[0]["body"]
        assert "예전 문장" not in same[0]["body"]
    finally:
        for name in os.listdir(folder):
            os.remove(os.path.join(folder, name))
        os.rmdir(folder)


def test_price_preview_html_lists_saved_rows():
    from price_increase_tab import _pi_price_preview_html

    html = _pi_price_preview_html(
        [
            {"품목명": "N2", "기존단가": 200, "인상적용단가": 220, "비고": "A"},
            {"품목명": "  ", "기존단가": 1, "인상적용단가": 1, "비고": ""},
        ],
        "2026-11-01",
    )
    assert "단가 조정 내용" in html
    assert "N2" in html
    assert "220" in html
    assert "2026년 11월 01일" in html
    assert _pi_price_preview_html([], "2026-11-01") == ""


def test_hold_mail_keeps_subject_and_recipient():
    import os
    import tempfile
    from datetime import date

    from price_increase_tab import (
        append_hold_mail,
        hold_mail_label,
        hold_recipient_label,
        load_hold_mail,
        _pi_restore_letter_kwargs,
    )

    folder = tempfile.mkdtemp()
    path = os.path.join(folder, "hold_mail.json")
    try:
        saved = append_hold_mail(
            {
                "to": "buyer@example.com",
                "cc": "",
                "subject": "제품가격 조정의 건",
                "body": "본문입니다.",
                "client": "경민산업",
                "attach_letter": True,
                "pdf_name": "공문.pdf",
                "letter_kwargs": {"send_date": date(2026, 11, 1), "title": "제품가격 조정의 건"},
                "send": True,
            },
            path,
        )
        rows = load_hold_mail(path)
        assert len(rows) == 1
        assert rows[0]["id"] == saved["id"]
        assert rows[0]["subject"] == "제품가격 조정의 건"
        assert rows[0]["to"] == "buyer@example.com"
        assert hold_recipient_label(rows[0]) == "경민산업 · buyer@example.com"
        assert "제품가격 조정의 건 · 경민산업 · buyer@example.com" in hold_mail_label(rows[0])
        assert hold_recipient_label({"to": "a@b.com"}, pd.DataFrame(columns=["거래처", "이메일"])) == "a@b.com"
        book = pd.DataFrame([{"거래처": "대진에스에이", "이메일": "djsy7615@naver.com"}])
        assert (
            hold_recipient_label({"to": "djsy7615@naver.com", "client": "일반메일"}, book)
            == "대진에스에이 · djsy7615@naver.com"
        )
        restored = _pi_restore_letter_kwargs(rows[0]["letter_kwargs"])
        assert restored["send_date"] == date(2026, 11, 1)
    finally:
        if os.path.isfile(path):
            os.remove(path)
        os.rmdir(folder)


def test_hold_tab_sits_between_bulk_and_history():
    src = open("price_increase_tab.py", encoding="utf-8").read()
    tabs = src.split('["개별 발송", "일괄 발송", "임시저장", "발송 이력"]', 1)
    assert len(tabs) == 2
    compose = src.split("def _render_pi_left_mail_compose", 1)[1].split("\ndef ", 1)[0]
    assert compose.index('"다시 임시저장" if editing_hold else "임시저장"') < compose.index(
        '"최종 발송"'
    )


def test_sent_history_opens_body_from_row():
    src = open("price_increase_tab.py", encoding="utf-8").read()
    hist = src.split("def _render_pi_sent_history", 1)[1].split("\ndef ", 1)[0]
    assert "거래처 검색" in hist
    assert "filter_sent_log_by_client" in hist
    assert 'on_select="rerun"' in hist
    assert 'selection_mode="single-row"' in hist
    assert "메일 본문" in hist
    assert src.count("body=body") >= 2
    assert "body=_with_mail_signature(body_b)" in src
