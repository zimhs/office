"""공문 개별발송: 연락처 이름이 거래처 검색에 포함되고 수신메일에 반영."""

import pandas as pd

from price_increase_tab import (
    _BODY_CELL_ORDER,
    _DEFAULT_LETTER_PARAS,
    _body_text_to_paras,
    _norm_client_list,
    _person_key,
    _pi_attach_letter_pdf,
    _plain_mail_intro,
    _strip_letter_body_from_kwargs,
    filter_letter_client_names,
    join_recipient_emails,
    list_clients_for_letter,
    list_clients_for_staff,
    list_mail_contact_names,
    lookup_email_with_meta,
    lookup_emails_for_clients,
)


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


def test_plain_mail_skips_letter_pdf():
    assert _pi_attach_letter_pdf(True) is True
    assert _pi_attach_letter_pdf(False) is False
    assert _plain_mail_intro("김도엽이사님") == "김도엽이사님 귀중\n\n"
    assert _plain_mail_intro("") == ""
