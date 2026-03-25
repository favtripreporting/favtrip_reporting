""" 
gmail_utils
======================================
Email utilities for sending Gmail messages with PDF attachments via the Gmail API.

This module provides small, focused helpers for sending two types of emails through
the Gmail API using a pre-authorized `gmail_svc` client (e.g., returned by
`googleapiclient.discovery.build("gmail", "v1", ...)`). It includes:

- `send_email(...)`: Low-level helper that accepts an `email.message.EmailMessage`,
  base64-url encodes it as required by Gmail, and dispatches it via
  `users.messages.send`.

- `email_manager_report(...)`: Composes and sends a standardized "Manager Report"
  email with a primary PDF attachment and a backup link. Supports optional CC.

- `email_order_report(...)`: Composes and sends an "Order Report" email for a
  given vendor or category key, including a primary PDF attachment and an optional
  "full order" PDF. Also includes links to a backing Google Sheet and supports CC.

The functions here intentionally perform minimal validation and assume that callers
supply valid addresses, attachments, and links. Authentication, token refresh, and
error handling policy (e.g., retries, backoff, alerting) should be implemented by
the caller.

---
Key Behaviors
-------------
- **MIME construction**: Uses Python's stdlib `email.message.EmailMessage` to build
  multipart emails with both plain-text and HTML alternatives, and PDF attachments.
- **Gmail API compliance**: Serializes the email to bytes and encodes it with
  URL-safe Base64 as required by Gmail's `users.messages.send` endpoint.
- **Idempotency**: Sending is not idempotent; calling functions repeatedly may
  result in duplicate emails. Callers should implement their own guardrails if
  needed (e.g., deduplication keys, sent-flagging).
- **Internationalization**: The functions do not localize content; callers can adapt
  the text if i18n is required.
- **HTML content**: Simple HTML bodies are included via `add_alternative(..., subtype="html")`.
  The HTML snippets intentionally avoid external assets for reliable delivery.

---
Functions
---------
send_email(gmail_svc, user, msg)
    Low-level send helper. Encodes the `EmailMessage` and dispatches via the Gmail API.

email_manager_report(gmail_svc, sender, to_list, cc_list, pdf_name, pdf_bytes, pdf_link, ts, location)
    Sends a standardized "Manager Report" email with a PDF attachment and a backup link.

email_order_report(
    gmail_svc,
    sender,
    to_list,
    cc_list,
    key,
    tag,
    ts,
    location,
    pdf_name,
    pdf_bytes,
    sheet_link,
    include_full_order=False,
    full_pdf_bytes=None,
    full_pdf_name=None,
)
    Sends an "Order Report" email targeted to a `{key}` team with a primary PDF,
    optional full-order PDF, and a link to the backing Google Sheet.

---
Parameters (Shared Concepts)
----------------------------
gmail_svc : Any
    An authenticated Gmail API service client (e.g., from `googleapiclient.discovery.build`).

sender : str
    The "From" email address to display in the message header. The authenticated
    Gmail account must be authorized to send from this address.

to_list : Iterable[str]
    Recipient email addresses for the `To` field. Must contain at least one valid address.

cc_list : Optional[Iterable[str]]
    Optional CC recipient addresses. If empty or `None`, the `Cc` header is omitted.

pdf_name : str
    Filename for the attached PDF (e.g., `"report_2026-03-21.pdf"`).

pdf_bytes : bytes
    Raw bytes of the primary PDF attachment.

ts : str
    A timestamp string suitable for inclusion in the subject (e.g., `"2026-03-21"` or
    `"2026-03-21 18:25"`).

location : str
    A human-readable location name included in the subject/body (e.g., store or site).

pdf_link : str
    (Manager Report) A backup URL users can access if attachments are blocked.

key : str
    (Order Report) An identifier for the receiving team or vendor (e.g., `"Dairy"`, `"VendorX"`).

tag : str
    (Order Report) A secondary descriptor (e.g., `"Weekly"`, `"Overstock"`, `"Emergency"`).

sheet_link : str
    (Order Report) URL to the backing Google Sheet with order details.

include_full_order : bool
    (Order Report) Whether to attach an additional "full order" PDF.

full_pdf_bytes : Optional[bytes]
    (Order Report) Raw bytes of the full order PDF (required when `include_full_order=True`).

full_pdf_name : Optional[str]
    (Order Report) Filename for the full order PDF (required when `include_full_order=True`).

user : str
    (send_email) Gmail user identifier for the API call. Typically `"me"` to refer
    to the authenticated account.

msg : EmailMessage
    (send_email) A fully-constructed email message to be sent.

---
Returns
-------
dict
    The Gmail API response payload from `users.messages.send()` (e.g., includes `id`, `threadId`).

---
Raises
------
googleapiclient.errors.HttpError
    If the Gmail API call fails (e.g., quota exceeded, invalid permissions, bad request).
ValueError / TypeError
    If provided inputs (addresses, bytes, filenames) are invalid (may be raised by stdlib or caller validations).

---
"""


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
        f"""
        <p>Hi team,</p>
        <p>Your manager report for store <b>{location}</b> is ready.</p>
        <p><a href='{pdf_link}'>Backup Link</a></p>
        <p>Attached: {pdf_name}</p>
        <p>—Automated</p>
        """,
        subtype="html",
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
        <p>Your order report for store <b>{location}</b> is ready.</p>
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


def email_error_report(
    gmail_svc,
    sender: str,
    to_list,
    cc_list,
    ts: str,
    pdf_name: str,
    pdf_bytes: bytes,
    sheet_link: str,
    vendor_price_book_link: str
    ):
    msg = EmailMessage()

    msg["Subject"] = f"Error Report – {ts}"
    msg["From"] = sender
    msg["To"] = ", ".join(to_list)

    if cc_list:
        msg["Cc"] = ", ".join(cc_list)

    msg.set_content(
        f"Hi team,\n"
        f"Some of the items you uploaded are not listed on the Vendor Price Book.\n"
        f"Google Sheet: {sheet_link}\n"
        f"Vendor Price Book: {vendor_price_book_link}"
        f"Attached: {pdf_name}\n"
        "—Automated"
    )

    msg.add_alternative(
        f"""
        <p>Hi team,</p>
        <p>Some of the items you uploaded are not listed on the Vendor Price Book.</p>
        <p><a href="{sheet_link}">Open Error Report in Google Sheets</a></p>
        <p><a href="{vendor_price_book_link}">Edit Vendor Price Book in Google Sheets</a></p>
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

    return send_email(gmail_svc, sender, msg)
