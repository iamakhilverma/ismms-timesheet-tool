#!/usr/bin/env python3
"""Outlook (macOS) backend: compose + send the timesheet email through the
locally signed-in Microsoft Outlook app via AppleScript.

This needs no Microsoft Graph permissions / admin consent -- it uses your already
authenticated Outlook session. Outlook has no scripted "delay delivery", so this
backend sends immediately (mode="send") or saves a ready-to-send draft to your
Drafts folder (mode="draft"). A saved draft syncs via Exchange to your phone /
other devices, so you can press Send from anywhere -- no scheduling needed.
"""
import os
import subprocess
import tempfile
from pathlib import Path

APP = "Microsoft Outlook"


def available() -> bool:
    return Path("/Applications/Microsoft Outlook.app").exists()


def _osa(script: str):
    r = subprocess.run(["osascript", "-e", script], capture_output=True, text=True)
    return r.returncode, (r.stdout or "").strip(), (r.stderr or "").strip()


def account_count():
    """Best-effort count of mail accounts Outlook exposes to automation.
    Returns an int, or None if it can't be determined (e.g. New Outlook hides them)."""
    code, out, _ = _osa(
        f'tell application "{APP}" to return (count of exchange accounts) + '
        f'(count of imap accounts) + (count of pop accounts)')
    return int(out) if code == 0 and out.lstrip("-").isdigit() else None


def sent_has(subject: str):
    """Best-effort count of Sent-Items messages with this exact subject (None if unknown)."""
    s = subject.replace("\\", "\\\\").replace('"', '\\"')
    code, out, _ = _osa(
        f'tell application "{APP}" to return count of (messages of sent items whose subject is "{s}")')
    return int(out) if code == 0 and out.isdigit() else None


def is_new_outlook() -> bool:
    r = subprocess.run(["defaults", "read", "com.microsoft.Outlook", "IsRunningNewOutlook"],
                       capture_output=True, text=True)
    return r.returncode == 0 and r.stdout.strip() == "1"


CLASSIC_HINT = (
    "Switch to Classic Outlook: in Outlook, click the 'New Outlook' toggle at the top-right "
    "to turn it OFF (or Outlook menu / Help menu > uncheck 'New Outlook'). Classic Outlook's "
    "automation can reach your account and Drafts. Then retry. (Nothing was sent.)")


def preflight():
    """Return None if the Outlook backend can run, else a clear error message.

    Note: we do NOT block on account_count()==0. New Outlook hides account
    enumeration from AppleScript yet still sends/saves fine, so blocking on the
    count would break the working path. Honesty comes from confirming in Sent
    Items / Drafts after the fact instead.
    """
    if not available():
        return "Microsoft Outlook is not installed in /Applications."
    if account_count() == 0 and not is_new_outlook():
        # Classic Outlook genuinely reports its accounts; 0 here means not signed in.
        return ("Outlook has no mail account signed in. Add your work account in Outlook, "
                "let it sync, then retry. (Nothing was sent.)")
    return None


def _q(s: str) -> str:
    """Quote a Python string as an AppleScript string literal."""
    s = str(s).replace("\\", "\\\\").replace('"', '\\"')
    s = s.replace("\r\n", "\n").replace("\r", "\n").replace("\n", '" & linefeed & "')
    return '"' + s + '"'


def _html_body(s: str) -> str:
    """Outlook's message `content` is HTML, which collapses plain newlines. Convert the
    plain-text body to HTML so line breaks (e.g. after 'Dear Team,') are preserved."""
    s = s.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    return s.replace("\n", "<br>")


def _script(cfg, pdf_path: Path, period: str, mode: str, show: bool, new_outlook: bool) -> str:
    e = cfg["email"]
    body = e["body"].format(period=period, name=cfg.get("name", ""))
    pdf = str(Path(pdf_path).resolve())
    lines = [
        f'tell application "{APP}"',
        f'  set m to make new outgoing message with properties {{subject:{_q(e["subject"])}, content:{_q(_html_body(body))}}}',
    ]
    for a in e["to"]:
        lines.append(f'  make new to recipient at m with properties {{email address:{{address:{_q(a)}}}}}')
    for a in e.get("cc", []):
        lines.append(f'  make new cc recipient at m with properties {{email address:{{address:{_q(a)}}}}}')
    lines.append(f'  make new attachment at m with properties {{file:(POSIX file {_q(pdf)})}}')
    if mode == "send":
        lines.append("  send m")
    elif new_outlook:
        # New Outlook can't script-save to Drafts, but it opens a fully-populated
        # compose window -> use its 'Schedule Send' (server-side) or save to Drafts.
        lines.append("  open m")
    else:
        # Classic Outlook: `make new outgoing message` already persists it to the
        # Drafts folder (which syncs). Do NOT call `save` (it errors -1701, wanting a
        # file path). Optionally open the window so you can review it.
        if show:
            lines.append("  try")
            lines.append("    open m")
            lines.append("  end try")
    lines.append('  return "ok"')
    lines.append("end tell")
    return "\n".join(lines)


def deliver(cfg, pdf_path, period, mode="send", show=True):
    """mode='send' sends immediately; mode='draft' saves a synced draft
    (show=True also opens its window)."""
    err = preflight()
    if err:
        raise RuntimeError(err)
    script = _script(cfg, pdf_path, period, mode, show, is_new_outlook())
    tmp = tempfile.NamedTemporaryFile("w", suffix=".applescript", delete=False)
    try:
        tmp.write(script)
        tmp.close()
        r = subprocess.run(["osascript", tmp.name], capture_output=True, text=True)
    finally:
        os.unlink(tmp.name)
    if r.returncode != 0:
        err = (r.stderr or r.stdout).strip()
        if "-1743" in err or "not allowed" in err.lower() or "not authoriz" in err.lower():
            raise RuntimeError(
                "macOS blocked controlling Outlook. Grant permission in System Settings > "
                "Privacy & Security > Automation (allow your terminal/Python to control Outlook), "
                "then retry. Original error: " + err)
        raise RuntimeError("Outlook AppleScript failed: " + err)
    return r.stdout.strip()
