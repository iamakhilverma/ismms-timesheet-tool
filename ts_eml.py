#!/usr/bin/env python3
"""Universal fallback backend: write the complete email as a standard .eml file
(recipients + body + PDF attachment). Needs NO account, NO admin, NO Outlook
sign-in -- it just produces a file you can open in Outlook (desktop or web) and
send, on any machine and in any state.
"""
from email.message import EmailMessage
from pathlib import Path


def build_eml(cfg, pdf_path, period, out_path):
    e = cfg["email"]
    msg = EmailMessage()
    if cfg.get("user_email"):
        msg["From"] = cfg["user_email"]
    msg["To"] = ", ".join(e["to"])
    if e.get("cc"):
        msg["Cc"] = ", ".join(e["cc"])
    msg["Subject"] = e["subject"]
    msg.set_content(e["body"].format(period=period, name=cfg.get("name", "")))
    pdf_path = Path(pdf_path)
    msg.add_attachment(pdf_path.read_bytes(), maintype="application",
                       subtype="pdf", filename=pdf_path.name)
    out_path = Path(out_path)
    out_path.write_bytes(bytes(msg))
    return out_path
