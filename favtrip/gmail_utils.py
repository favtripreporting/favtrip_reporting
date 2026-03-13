from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {location} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)
    if cc_list:
        msg["Cc"] = ", ".join(cc_list)
    msg.set_content(f"Hi team,\nAttached is the Manager Report ({location}).\nBackup link: {pdf_link}\n—Automated")
    msg.add_alternative(
        f"<p>Hi team,</p><p>Manager Report ({location})</p>"
        f"<a href='{pdf_link}'>Backup Link</a>", subtype="html"
    )
    msg.add_attachment(pdf_bytes, maintype="application", subtype="pdf", filename=pdf_name)
    return send_email(gmail_svc, sender, msg)


def email_order_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    key: str,
    tag: str,
    ts: str,
    location: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    include_full_order: bool = False,
    full_pdf_bytes: bytes | None = None,
    full_pdf_name: str | None = None,
):
    msg = EmailMessage()

    msg["Subject"] = f"Order Report – {location} – {tag} – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi {key} team,\n"
        f"Your order report for {location} - {tag} is ready.\n"
        f"Google Sheet: {sheet_link}\n"
        f"Attached: {pdf_name}\n"
        "—Automated"
    )

    msg.add_alternative(
        f"""
        <p>Hi {key} team,</p>
        <p>Your order report for store <b>{store}</b> is ready.</p>
        <p><a href="{sheet_link}">Open Google Sheet</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Automated</p>
        """,
        subtype="html",
    )

    msg.add_attachment(
        pdf_bytes,
        maintype="application",
        subtype="pdf",
        filename=pdf_name,
    )

    if include_full_order and full_pdf_bytes:
        msg.add_attachment(
            full_pdf_bytes,
            maintype="application",
            subtype="pdf",
            filename=full_pdf_name,
        )

    return send_email(gmail_svc, sender, msg)
