from src.models import Contact


def test_contact_recipient_uses_lookup_email_when_available() -> None:
    contact = Contact(
        values=["12345", "Some value"],
        row_number=2,
        headers=["id_code", "note"],
        lookup_map={"12345": "user@example.com"},
    )

    assert contact.recipient == "user@example.com"


def test_contact_recipient_falls_back_to_first_column_when_lookup_missing() -> None:
    contact = Contact(
        values=["12345", "Some value"],
        row_number=2,
        headers=["id_code", "note"],
        lookup_map={"99999": "user@example.com"},
    )

    assert contact.recipient == "12345"
