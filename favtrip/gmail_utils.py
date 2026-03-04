from __future__ import annotations
import base64
from email.message import EmailMessage


def send_email(gmail_svc, user: str, msg: EmailMessage):
    raw = base64.urlsafe_b64encode(msg.as_bytes()).decode()
    return gmail_svc.users().messages().send(userId=user, body={"raw": raw}).execute()


def email_manager_report(gmail_svc, sender: str, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location):
    msg = EmailMessage()
    msg["Subject"] = f"Manager Report – {ts} – {location}"
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
