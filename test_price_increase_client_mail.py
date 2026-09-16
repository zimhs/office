"""공문 개별발송: 연락처 이름이 거래처 검색에 포함되고 수신메일에 반영."""

import pandas as pd

from price_increase_tab import (
    _apply_amount_to_items,
    _apply_pct_to_items,
    _increase_mode_is_pct,
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
    join_recipient_emails,
    merge_keep_manual_emails,
    remove_email_addr,
    list_clients_for_letter,
    list_clients_for_staff,
    list_mail_contact_names,
    lookup_email_with_meta,
    lookup_emails_for_clients,
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
