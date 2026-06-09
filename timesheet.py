#!/usr/bin/env python3
"""
Weekly Time & Attendance timesheet tool.

Builds the filled PDF for a Sun->Sat payroll week and either sends it now or
schedules it for later via Microsoft 365. Scheduling uses Exchange's own
server-side deferred delivery (the "Delay Delivery" mechanism): you approve one
MFA push, and the mail server sends it at the chosen time -- your Mac does not
need to be on. No background agents, no local scheduler.

Examples
--------
  ./timesheet login                       one-time sign-in (approve MFA)
  ./timesheet schedule                     schedule THIS week's sheet for Fri 2:00 PM
  ./timesheet schedule --holiday fri        ...mark Friday as holiday (H)
  ./timesheet schedule --pto wed            ...Wed = PTO (scheduled, paid 7.5h)
  ./timesheet schedule --pto wed:unpaid     ...Wed = PTO unpaid (drops weekly total)
  ./timesheet schedule --pto wed:ptou       ...Wed = PTO unscheduled (PTOU)
  ./timesheet schedule --at "fri 9am"       ...different send time
  ./timesheet schedule --week 2026-07-12    ...a different week
  ./timesheet send                          build + send right now
  ./timesheet build  [flags]                just build the PDF, no email
  ./timesheet preview [flags]               build + print the email, send nothing
  ./timesheet whoami | logout | doctor
"""
import argparse
import base64
import datetime as dt
import json
import sys
from pathlib import Path

HERE = Path(__file__).resolve().parent

try:
    import requests
    import ts_pdf
    import ts_outlook
    import ts_eml
except ImportError:
    sys.stderr.write(
        "Missing dependencies. Run the one-time setup first:\n"
        f"    cd {HERE} && ./setup.sh\n"
        "then use ./timesheet ...\n")
    sys.exit(1)

CONFIG_PATH = HERE / "config.json"
PROFILE_PATH = HERE / "profile.json"      # YOUR personal info (git-ignored)
TEMPLATE_PATH = HERE / "template.pdf"     # YOUR blank form (git-ignored)
TOKEN_CACHE = HERE / ".token_cache.json"
BACKEND_FILE = HERE / ".backend"          # per-machine choice: "outlook" or "graph"
OUTPUT_DIR = HERE / "output"
GRAPH = "https://graph.microsoft.com/v1.0"
# Minimal scope for sending now -- usually user-consentable (no admin approval).
SEND_SCOPES = ["Mail.Send", "User.Read"]
# Scheduling (server-side deferred send) also needs to create/read a draft, which
# requires the broader Mail.ReadWrite -- this is the one a tenant may gate on admin.
SCHEDULE_SCOPES = ["Mail.Send", "Mail.ReadWrite", "User.Read"]
# PidTagDeferredSendTime (PR_DEFERRED_SEND_TIME) -> Exchange server-side delay.
DEFERRED_SEND_PROP = "SystemTime 0x3FEF"

WEEKDAYS = {"sun": 6, "mon": 0, "tue": 1, "wed": 2, "thu": 3, "fri": 4, "sat": 5}
WEEKDAY_FULL = {"sunday", "monday", "tuesday", "wednesday", "thursday", "friday", "saturday"}


# ----------------------------- config ---------------------------------------
def _deep_merge(base, over):
    out = dict(base)
    for k, v in over.items():
        if isinstance(v, dict) and isinstance(out.get(k), dict):
            out[k] = _deep_merge(out[k], v)
        else:
            out[k] = v
    return out


def load_config():
    with open(CONFIG_PATH) as fh:
        cfg = json.load(fh)
    if PROFILE_PATH.exists():               # personal info overrides the shared defaults
        with open(PROFILE_PATH) as fh:
            cfg = _deep_merge(cfg, json.load(fh))
    return cfg


def _missing_profile_fields(cfg):
    miss = [k for k in ("name", "life_num", "user_email") if not str(cfg.get(k, "")).strip()]
    if not cfg.get("email", {}).get("to"):
        miss.append("recipients (to)")
    return miss


def require_profile(cfg):
    miss = _missing_profile_fields(cfg)
    if miss:
        raise RuntimeError("Your profile isn't set up yet (missing: " + ", ".join(miss) +
                           ").\nRun:  ./timesheet setup   (or copy profile.example.json to profile.json)")


def require_template():
    if not TEMPLATE_PATH.exists():
        raise RuntimeError(
            "No template.pdf found. Put your blank fillable 'Daily Attendance Record' PDF here as\n"
            f"  {TEMPLATE_PATH}\n(It stays local -- git-ignored so your signature/name never leave your Mac. See the README.)")


def setup_profile():
    """Prompt for personal info and save it to profile.json (git-ignored)."""
    existing = {}
    if PROFILE_PATH.exists():
        existing = json.load(open(PROFILE_PATH))
    em = existing.get("email", {})

    def ask(label, current):
        cur = f" [{current}]" if current else ""
        ans = input(f"  {label}{cur}: ").strip()
        return ans or (current or "")

    print("\nLet's set up your profile (saved locally to profile.json, never committed):")
    name = ask("Your full name (as it appears on the sheet)", existing.get("name"))
    life = ask("Your Life # (employee ID)", existing.get("life_num"))
    email = ask("Your work email", existing.get("user_email"))
    to = ask("Send TO (comma-separated)", ", ".join(em.get("to", ["GGSTimeAttendance@mssm.edu"])))
    cc = ask("Cc (your supervisor, comma-separated; blank for none)", ", ".join(em.get("cc", [])))
    profile = {
        "name": name, "life_num": life, "user_email": email,
        "email": {
            "to": [a.strip() for a in to.split(",") if a.strip()],
            "cc": [a.strip() for a in cc.split(",") if a.strip()],
        },
        "holidays": existing.get("holidays", []),
        "pto": existing.get("pto", []),
    }
    PROFILE_PATH.write_text(json.dumps(profile, indent=2) + "\n")
    print(f"Saved your profile to {PROFILE_PATH.name} (git-ignored).")
    return profile


# ----------------------------- backend (per machine) ------------------------
def choose_backend(persist=True):
    """Interactively pick how THIS machine sends mail; remembers the choice."""
    print("\nHow should this machine send the timesheet email?")
    print("  1) outlook  - through the Microsoft Outlook app you're signed into")
    print("                (no admin approval needed; sends from this Mac)")
    print("  2) graph    - via Microsoft 365 API (enables true server-side scheduling,")
    print("                but your org must approve the app's mail permission)")
    ans = input("Choose [1/2] (default 1=outlook): ").strip().lower()
    backend = "graph" if ans in ("2", "graph", "g") else "outlook"
    if persist:
        BACKEND_FILE.write_text(backend + "\n")
        print(f"Saved: this machine will use the '{backend}' method (edit/remove {BACKEND_FILE.name} to change).")
    return backend


def resolve_backend(args, cfg, interactive=True):
    via = getattr(args, "via", None)
    if via:
        return via
    if BACKEND_FILE.exists():
        v = BACKEND_FILE.read_text().strip()
        if v in ("outlook", "graph"):
            return v
    if cfg.get("backend") in ("outlook", "graph"):
        return cfg["backend"]
    if interactive and sys.stdin.isatty():
        return choose_backend()
    # sensible default: Outlook if installed (no admin needed), else graph
    return "outlook" if ts_outlook.available() else "graph"


# ----------------------------- dates ----------------------------------------
def week_sunday(d: dt.date) -> dt.date:
    return d - dt.timedelta(days=(d.weekday() + 1) % 7)


def resolve_week(week_arg):
    """Resolve --week to the Sunday that starts the target week.

    Accepts: nothing/'this' = current week; 'next'/'last'; '+N'/'-N' (weeks from now,
    optional trailing 'w'); or any date in the week as YYYY-MM-DD.
    """
    import re
    if not week_arg:
        return week_sunday(dt.date.today())
    w = week_arg.strip().lower()
    this = week_sunday(dt.date.today())
    if w in ("this", "current", "now"):
        return this
    if w in ("next",):
        return this + dt.timedelta(days=7)
    if w in ("last", "prev", "previous"):
        return this - dt.timedelta(days=7)
    if re.fullmatch(r"[+-]\d+w?", w):
        return this + dt.timedelta(days=7 * int(w.rstrip("w")))
    return week_sunday(dt.date.fromisoformat(week_arg))


def resolve_day(token: str, week_start: dt.date) -> dt.date:
    """Map 'fri' / '19' / '2026-06-19' to a date inside the Sun..Sat week."""
    week = [week_start + dt.timedelta(days=i) for i in range(7)]
    t = token.strip().lower()
    if t[:3] in WEEKDAYS and (len(t) == 3 or t in WEEKDAY_FULL):
        return week[(WEEKDAYS[t[:3]] + 1) % 7]
    if t.isdigit():
        for day in week:
            if day.day == int(t):
                return day
        raise ValueError(f"day-of-month {t} is not in week {week[0]:%m/%d}-{week[-1]:%m/%d}")
    day = dt.date.fromisoformat(token)
    if not (week[0] <= day <= week[-1]):
        raise ValueError(f"{token} is not in week {week[0]:%m/%d}-{week[-1]:%m/%d}")
    return day


def parse_clock(s: str) -> dt.time:
    """'2pm' / '2:30pm' / '14:00' / '9am' -> time."""
    s = s.strip().lower().replace(" ", "")
    ampm = None
    if s.endswith("am") or s.endswith("pm"):
        ampm, s = s[-2:], s[:-2]
    h, m = (s.split(":") + ["0"])[:2] if ":" in s else (s, "0")
    h, m = int(h), int(m)
    if ampm == "pm" and h < 12:
        h += 12
    if ampm == "am" and h == 12:
        h = 0
    return dt.time(h, m)


def parse_when(when, week_start, cfg) -> dt.datetime:
    """Return a timezone-aware local datetime for the scheduled send.

    Default: Friday of the target week at config.default_send_time.
    `when` may be 'fri 2pm', '9am', 'wed 14:00', or 'YYYY-MM-DD HH:MM'.
    """
    local_tz = dt.datetime.now().astimezone().tzinfo
    day = week_start + dt.timedelta(days=5)                 # Friday
    clock = parse_clock(cfg.get("default_send_time", "14:00"))
    if when:
        w = when.strip()
        try:                                               # full datetime?
            return dt.datetime.fromisoformat(w).replace(tzinfo=local_tz)
        except ValueError:
            pass
        for tok in w.split():
            try:
                clock = parse_clock(tok)
            except (ValueError, IndexError):
                day = resolve_day(tok, week_start)
    return dt.datetime.combine(day, clock).replace(tzinfo=local_tz)


# ----------------------------- week plan ------------------------------------
def _expand_iso(item: str):
    """'2026-06-19' or '2026-08-03..2026-08-15' -> list of dates (inclusive)."""
    if ".." in item:
        a, b = (s.strip() for s in item.split("..", 1))
        a, b = dt.date.fromisoformat(a), dt.date.fromisoformat(b)
        if b < a:
            a, b = b, a
        return [a + dt.timedelta(days=i) for i in range((b - a).days + 1)]
    return [dt.date.fromisoformat(item)]


def _expand_token(token: str, week_start: dt.date):
    """A --holiday/--pto day token -> dates. Accepts weekday (mon..fri), a
    day-of-month number, an ISO date, or an ISO range A..B."""
    t = token.strip()
    if ".." in t:
        return _expand_iso(t)
    if "-" in t and t[:4].isdigit():      # ISO date
        return [dt.date.fromisoformat(t)]
    return [resolve_day(t, week_start)]   # weekday name / day number


def _parse_pto_opts(parts):
    code, paid = "PTOS", True
    for opt in parts:
        o = opt.lower()
        if o in ("ptos", "ptou"):
            code = o.upper()
        elif o in ("unpaid", "u"):
            paid = False
        elif o in ("paid", "p"):
            paid = True
        else:
            raise ValueError(f"unknown --pto option {opt!r} (use ptos/ptou/paid/unpaid)")
    return code, paid


def build_overrides(week_start, holiday_tokens, pto_tokens, cfg):
    """Combine config holidays/PTO + CLI flags into ({date -> override}, skipped_weekends).

    Holidays/PTO that fall on a weekend are skipped (weekends carry no hours);
    ranges therefore only mark the Mon-Fri days inside them.
    """
    weekset = {week_start + dt.timedelta(days=i) for i in range(7)}
    overrides, skipped = {}, []

    def put(d, ov):
        if d not in weekset:
            return
        if d.weekday() >= 5:              # Sat/Sun -> no hours
            skipped.append(d)
            return
        overrides[d] = ov

    for item in cfg.get("holidays", []):          # config holidays (dates/ranges)
        for d in _expand_iso(item):
            put(d, {"kind": "holiday"})
    for item in cfg.get("pto", []):               # config PTO
        code = item.get("code", "PTOS")
        paid = bool(item.get("paid", True))
        if item.get("from") and item.get("to"):
            days = _expand_iso(f"{item['from']}..{item['to']}")
        elif item.get("date"):
            days = _expand_iso(item["date"])
        else:
            continue
        for d in days:
            put(d, {"kind": "pto", "code": code, "paid": paid})
    for tok in holiday_tokens or []:              # CLI holidays
        for d in _expand_token(tok, week_start):
            put(d, {"kind": "holiday"})
    for tok in pto_tokens or []:                  # CLI PTO  day[:opts]
        parts = tok.split(":")
        code, paid = _parse_pto_opts(parts[1:])
        for d in _expand_token(parts[0], week_start):
            put(d, {"kind": "pto", "code": code, "paid": paid})
    return overrides, skipped


def describe(overrides, cfg):
    lines = []
    for d in sorted(overrides):
        o = overrides[d]
        if o["kind"] == "holiday":
            lines.append(f"  - {d:%a %m/%d}: HOLIDAY ({cfg['holiday_code']})")
        else:
            lines.append(f"  - {d:%a %m/%d}: PTO ({o['code']} {'paid' if o.get('paid', True) else 'UNPAID'})")
    return lines or ["  - normal week (Mon-Fri 9:00-5:00)"]


def email_body(cfg, period):
    return cfg["email"]["body"].format(period=period, name=cfg.get("name", ""))


def build_pdf(cfg, week_start, holiday_tokens, pto_tokens, announce=True):
    require_template()
    require_profile(cfg)
    overrides, skipped = build_overrides(week_start, holiday_tokens, pto_tokens, cfg)
    week_end = week_start + dt.timedelta(days=6)
    period = f"{week_start:%m/%d} to {week_end:%m/%d}"
    period_file = f"{week_start:%m-%d} to {week_end:%m-%d}"
    fname = cfg["email"]["attachment_name"].format(period=period, period_file=period_file)
    out = OUTPUT_DIR / fname
    fields, _, _ = ts_pdf.fill(week_start, out, overrides, cfg)
    if announce:
        print(f"Week:   {week_start:%a %m/%d/%Y} -> {week_end:%a %m/%d/%Y}")
        print(f"Period: {period}   Total hours: {fields['Total Hours']}")
        for line in describe(overrides, cfg):
            print(line)
        if skipped:
            print(f"  (ignored weekend day(s): {', '.join(f'{d:%a %m/%d}' for d in sorted(set(skipped)))})")
        print(f"PDF:    {out}")
    return out, period


# ----------------------------- auth (MSAL) ----------------------------------
def _msal_app(cfg):
    import msal
    cache = msal.SerializableTokenCache()
    if TOKEN_CACHE.exists():
        cache.deserialize(TOKEN_CACHE.read_text())
    app = msal.PublicClientApplication(
        cfg["azure_client_id"],
        authority=f"https://login.microsoftonline.com/{cfg['tenant']}",
        token_cache=cache)
    return app, cache


def get_token(cfg, scopes, interactive=True, use_device=False):
    app, cache = _msal_app(cfg)
    result = None
    for acct in app.get_accounts():
        result = app.acquire_token_silent(scopes, account=acct)
        if result:
            break
    if not result and interactive:
        hint = cfg.get("user_email") or None
        if use_device:
            flow = app.initiate_device_flow(scopes=scopes)
            if "user_code" not in flow:
                raise RuntimeError(f"device flow failed: {flow.get('error_description', flow)}")
            print("\n" + "=" * 64 + f"\n{flow['message']}\n" + "=" * 64 + "\n", flush=True)
            result = app.acquire_token_by_device_flow(flow)
        else:
            # Interactive browser (auth-code+PKCE). Runs in the real browser on the
            # managed device, which passes Conditional Access policies that block the
            # device-code flow.
            print("Opening your browser to sign in -- complete it and approve MFA there...", flush=True)
            result = app.acquire_token_interactive(scopes, prompt="select_account", login_hint=hint)
    if cache.has_state_changed:
        TOKEN_CACHE.write_text(cache.serialize())
        try:
            TOKEN_CACHE.chmod(0o600)
        except Exception:
            pass
    if not result or "access_token" not in result:
        raise RuntimeError((result or {}).get(
            "error_description", "not signed in -- run: ./timesheet login"))
    return result["access_token"]


def _headers(token, json_body=False):
    h = {"Authorization": f"Bearer {token}"}
    if json_body:
        h["Content-Type"] = "application/json"
    return h


def graph_get(token, path):
    r = requests.get(GRAPH + path, headers=_headers(token))
    r.raise_for_status()
    return r.json()


# ----------------------------- message build / send -------------------------
def _message(cfg, pdf_path: Path, period: str, deferred_utc: str = None):
    e = cfg["email"]
    with open(pdf_path, "rb") as fh:
        content_b64 = base64.b64encode(fh.read()).decode()
    msg = {
        "subject": e["subject"],
        "body": {"contentType": "Text", "content": email_body(cfg, period)},
        "toRecipients": [{"emailAddress": {"address": a}} for a in e["to"]],
        "ccRecipients": [{"emailAddress": {"address": a}} for a in e.get("cc", [])],
        "attachments": [{
            "@odata.type": "#microsoft.graph.fileAttachment",
            "name": pdf_path.name,
            "contentType": "application/pdf",
            "contentBytes": content_b64,
        }],
    }
    if deferred_utc:
        msg["singleValueExtendedProperties"] = [
            {"id": DEFERRED_SEND_PROP, "value": deferred_utc}]
    return msg


def send_now(token, cfg, pdf_path, period):
    r = requests.post(GRAPH + "/me/sendMail", headers=_headers(token, True),
                      data=json.dumps({"message": _message(cfg, pdf_path, period),
                                       "saveToSentItems": True}))
    if r.status_code not in (200, 202):
        raise RuntimeError(f"sendMail failed [{r.status_code}]: {r.text}")


def schedule_send(token, cfg, pdf_path, period, when_local: dt.datetime):
    """Create a draft carrying the deferred-send time, then submit it so
    Exchange delivers it server-side at `when_local`."""
    deferred_utc = when_local.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")
    r = requests.post(GRAPH + "/me/messages", headers=_headers(token, True),
                      data=json.dumps(_message(cfg, pdf_path, period, deferred_utc)))
    if r.status_code not in (200, 201):
        raise RuntimeError(f"create draft failed [{r.status_code}]: {r.text}")
    msg_id = r.json()["id"]
    r = requests.post(GRAPH + f"/me/messages/{msg_id}/send", headers=_headers(token))
    if r.status_code not in (200, 202):
        raise RuntimeError(f"submit failed [{r.status_code}]: {r.text}")


# ----------------------------- commands -------------------------------------
def _print_email(cfg, period):
    e = cfg["email"]
    print("\nEmail:")
    print(f"  To:      {', '.join(e['to'])}")
    print(f"  Cc:      {', '.join(e.get('cc', []))}")
    print(f"  Subject: {e['subject']}")
    print("  Body:")
    for line in email_body(cfg, period).splitlines():
        print(f"    {line}")


def cmd_build(args, cfg):
    build_pdf(cfg, resolve_week(args.week), args.holiday, args.pto)


def cmd_preview(args, cfg):
    _, period = build_pdf(cfg, resolve_week(args.week), args.holiday, args.pto)
    _print_email(cfg, period)
    print("\n[preview] nothing sent.")


def cmd_send(args, cfg):
    out, period = build_pdf(cfg, resolve_week(args.week), args.holiday, args.pto)
    _print_email(cfg, period)
    backend = resolve_backend(args, cfg)
    if backend == "outlook":
        print("\nSending now via the Outlook app ...")
        ts_outlook.deliver(cfg, out, period, mode="send")
        print("Sent (handed to Outlook).")
        return
    token = get_token(cfg, SEND_SCOPES)
    who = graph_get(token, "/me").get("userPrincipalName", "?")
    print(f"\nSending now as {who} (Graph) ...")
    send_now(token, cfg, out, period)
    print("Sent.")


def cmd_draft(args, cfg):
    """Build the PDF and stage the email in Outlook (compose window or saved draft)."""
    out, period = build_pdf(cfg, resolve_week(args.week), args.holiday, args.pto)
    _print_email(cfg, period)
    ts_outlook.deliver(cfg, out, period, mode="draft")
    if ts_outlook.is_new_outlook():
        print("\nOpened a fully-filled compose window in Outlook (To/Cc/Subject/body + PDF). Two options:")
        print("  - Friday-2PM auto-send: click the dropdown arrow next to Send > 'Send Later' /")
        print("    'Schedule Send', pick Friday 2:00 PM. Exchange sends it then (Mac can be off).")
        print("  - Or save to Drafts (Cmd+S): it syncs to your phone -- tap Send anytime.")
    else:
        print("\nSaved a ready-to-send draft to Outlook Drafts. It syncs to your phone -- tap Send Friday.")


def cmd_eml(args, cfg):
    """Write the complete email as a .eml file (PDF attached). No account/admin needed.
    Open it in Outlook (desktop or web) and press Send."""
    import subprocess
    week_start = resolve_week(args.week)
    out, period = build_pdf(cfg, week_start, args.holiday, args.pto)
    eml_path = OUTPUT_DIR / (f"Time & Attendance email ({week_start:%m-%d}"
                             f" to {week_start + dt.timedelta(days=6):%m-%d}).eml")
    ts_eml.build_eml(cfg, out, period, eml_path)
    _print_email(cfg, period)
    print(f"\nWrote ready-to-send email: {eml_path}")
    if not args.no_open:
        subprocess.run(["open", str(eml_path)], check=False)
        print("Opened it in your default mail app -- review and press Send.")
    else:
        print("Open it in Outlook (desktop or web) and press Send.")


def cmd_setup(args, cfg):
    setup_profile()
    if not TEMPLATE_PATH.exists():
        print("\nIMPORTANT: place your blank fillable 'Daily Attendance Record' PDF here as")
        print(f"  {TEMPLATE_PATH.name}   (stays local; git-ignored). See the README.")
    backend = choose_backend()
    if backend == "graph":
        print("\nNow sign in for Graph:")
        cmd_login(args, cfg)
    else:
        print("\nOutlook method selected. Make sure the Outlook app is signed into your work account.")
        print("Then test it with:  ./timesheet verify   (sends only to you).")


def _weeks_to_schedule(start_week, through, weeks):
    if through:
        end_week = week_sunday(dt.date.fromisoformat(through))
        if end_week < start_week:
            raise ValueError("--through is before the target week")
        out, w = [], start_week
        while w <= end_week:
            out.append(w)
            w += dt.timedelta(days=7)
        return out
    if weeks and weeks > 1:
        return [start_week + dt.timedelta(days=7 * i) for i in range(weeks)]
    return [start_week]


def cmd_schedule(args, cfg):
    backend = resolve_backend(args, cfg)
    week_starts = _weeks_to_schedule(resolve_week(args.week), args.through, args.weeks)
    if backend == "outlook":
        single = len(week_starts) == 1
        for ws in week_starts:
            out, period = build_pdf(cfg, ws, args.holiday, args.pto, announce=single)
            ts_outlook.deliver(cfg, out, period, mode="draft", show=single)
            if not single:
                print(f"  draft saved: {period}")
        n = len(week_starts)
        if ts_outlook.is_new_outlook():
            print(f"\nOpened {n} fully-filled compose window{'s' if n > 1 else ''} in Outlook (nothing sent).")
            print("For each: use 'Schedule Send' to set Friday 2 PM (server-side), or save to Drafts")
            print("(Cmd+S) to send from your phone.")
        else:
            print(f"\nSaved {n} ready-to-send draft{'s' if n > 1 else ''} to Outlook Drafts (nothing sent).")
            print("They sync to your phone -- open Outlook anywhere and press Send (or Schedule Send 2 PM).")
        return
    now = dt.datetime.now().astimezone()
    single = len(week_starts) == 1
    plan = []
    for ws in week_starts:
        out, period = build_pdf(cfg, ws, args.holiday, args.pto, announce=single)
        plan.append((out, period, parse_when(args.at, ws, cfg)))
    if single:
        _print_email(cfg, plan[0][1])
        print(f"\nDeliver at: {plan[0][2]:%a %m/%d/%Y %I:%M %p %Z} (server-side)")
        if plan[0][2] <= now:
            print("Warning: that time is in the past; Exchange will send it ~immediately.")
    else:
        print(f"\nScheduling {len(week_starts)} weekly emails:")
        for out, period, when in plan:
            print(f"  {period}  ->  deliver {when:%a %m/%d %I:%M %p}"
                  + ("   (PAST -> sends ~now)" if when <= now else ""))
    token = get_token(cfg, SCHEDULE_SCOPES)
    who = graph_get(token, "/me").get("userPrincipalName", "?")
    print(f"Scheduling as {who} ...")
    for out, period, when in plan:
        schedule_send(token, cfg, out, period, when)
        print(f"  scheduled  {period}  for  {when:%a %m/%d %I:%M %p}")
    print("Done. Exchange delivers each at its time (your Mac can be off).")
    print("To cancel/edit before then: Outlook on the web -> Drafts.")


def _self_test_cfg(cfg, period):
    """A config clone addressed ONLY to you, for safe self-tests."""
    me = cfg.get("user_email")
    if not me:
        raise RuntimeError("set user_email in config.json to run a self-test")
    t = json.loads(json.dumps(cfg))
    t["email"]["to"] = [me]
    t["email"]["cc"] = []
    t["email"]["subject"] = "[TEST] Time & Attendance tool self-test"
    t["email"]["body"] = ("This is an automated self-test of the timesheet tool. "
                          f"It is addressed only to you. Sample period: {period}.")
    return t, me


def cmd_verify(args, cfg):
    """Prove the pipeline works by sending a test ONLY to yourself. Never emails the
    payroll recipients, never deletes anything."""
    week_start = resolve_week(None)
    out, period = build_pdf(cfg, week_start, [], [], announce=False)
    if resolve_backend(args, cfg) == "outlook":
        import time
        tcfg, me = _self_test_cfg(cfg, period)
        err = ts_outlook.preflight()
        if err:
            raise RuntimeError(err)
        ts_outlook.deliver(tcfg, out, period, mode="send")
        found = None
        for _ in range(6):                       # confirm it really left (best-effort)
            found = ts_outlook.sent_has(tcfg["email"]["subject"])
            if found:
                break
            time.sleep(5)
        if found:
            print(f"CONFIRMED: test email sent to {me} (found in Outlook Sent Items). Check your inbox.")
        else:
            print(f"Sent to {me}. (New Outlook hides Sent Items from automation so I can't auto-confirm,")
            print(" but the send goes through -- check your inbox to be sure.)")
        return
    import time
    token = get_token(cfg, SCHEDULE_SCOPES)
    me = graph_get(token, "/me")
    addr = args.to or me.get("mail") or me.get("userPrincipalName")
    now = dt.datetime.now().astimezone()
    when_local = now + dt.timedelta(minutes=args.in_minutes)
    marker = when_local.strftime("%H%M%S")
    subject = f"[TEST {marker}] T&A deferred-send check"
    deferred_utc = when_local.astimezone(dt.timezone.utc).strftime("%Y-%m-%dT%H:%M:%SZ")

    with open(out, "rb") as fh:
        content_b64 = base64.b64encode(fh.read()).decode()
    msg = {
        "subject": subject,
        "body": {"contentType": "Text",
                 "content": f"Automated test of server-side scheduled send.\n"
                            f"Submitted ~{now:%H:%M:%S}, requested delivery {when_local:%H:%M:%S %Z}."},
        "toRecipients": [{"emailAddress": {"address": addr}}],
        "attachments": [{"@odata.type": "#microsoft.graph.fileAttachment",
                         "name": out.name, "contentType": "application/pdf",
                         "contentBytes": content_b64}],
        "singleValueExtendedProperties": [{"id": DEFERRED_SEND_PROP, "value": deferred_utc}],
    }
    r = requests.post(GRAPH + "/me/messages", headers=_headers(token, True), data=json.dumps(msg))
    if r.status_code not in (200, 201):
        raise RuntimeError(f"create draft failed [{r.status_code}]: {r.text}")
    mid = r.json()["id"]
    r = requests.post(GRAPH + f"/me/messages/{mid}/send", headers=_headers(token))
    if r.status_code not in (200, 202):
        raise RuntimeError(f"submit failed [{r.status_code}]: {r.text}")
    print(f"Submitted test to {addr} at {now:%H:%M:%S}, requested delivery {when_local:%H:%M:%S %Z}.")
    if args.no_wait:
        print(f"Check your inbox at ~{when_local:%H:%M}; subject: {subject}")
        return

    deadline = when_local + dt.timedelta(minutes=3)
    print("Watching inbox (Ctrl-C to stop)...", flush=True)
    while dt.datetime.now().astimezone() < deadline:
        q = ("/me/mailFolders/Inbox/messages?$top=20&$orderby=receivedDateTime desc"
             "&$select=subject,receivedDateTime")
        for m in graph_get(token, q).get("value", []):
            if m.get("subject") == subject:
                recv = dt.datetime.fromisoformat(m["receivedDateTime"].replace("Z", "+00:00")).astimezone()
                held = (recv - now).total_seconds()
                print(f"Delivered at {recv:%H:%M:%S %Z} (held ~{held/60:.1f} min after submit).")
                if held >= (args.in_minutes * 60) - 60:
                    print("VERIFIED: Exchange honored the deferred-send time -- scheduling works.")
                else:
                    print("NOTE: it arrived sooner than requested; this tenant may not defer server-side.")
                return
        time.sleep(15)
    print("Did not see the test arrive within the window. Check Inbox / Drafts manually,")
    print("and your Junk folder. If it never arrives, the tenant may block this flow.")


def cmd_login(args, cfg):
    scopes = SCHEDULE_SCOPES if getattr(args, "full", False) else SEND_SCOPES
    me = graph_get(get_token(cfg, scopes, True, use_device=getattr(args, "device", False)), "/me")
    print(f"Signed in as {me.get('displayName')} <{me.get('userPrincipalName')}>")
    print(f"Granted: {', '.join(scopes)}"
          + ("" if getattr(args, "full", False) else
             "\n(For server-side scheduling, run: ./timesheet login --full)"))


def cmd_whoami(args, cfg):
    try:
        me = graph_get(get_token(cfg, False), "/me")
    except Exception as ex:
        print(f"Not signed in ({ex}).")
        return
    print(f"{me.get('displayName')} <{me.get('userPrincipalName')}>")


def cmd_logout(args, cfg):
    if TOKEN_CACHE.exists():
        TOKEN_CACHE.unlink()
        print("Cleared cached login.")
    else:
        print("No cached login.")


def cmd_doctor(args, cfg):
    ok = True

    def check(label, good, detail=""):
        nonlocal ok
        ok = ok and good
        print(f"  [{'OK ' if good else 'XX '}] {label}{(' - ' + detail) if detail else ''}")

    print("timesheet doctor")
    check(f"Python {sys.version_info.major}.{sys.version_info.minor}", sys.version_info >= (3, 9),
          "" if sys.version_info >= (3, 9) else "need 3.9+")
    check("running inside .venv", str(HERE / ".venv") in sys.executable,
          sys.executable)
    for mod in ("pypdf", "msal", "requests"):
        try:
            m = __import__(mod)
            check(f"module {mod}", True, getattr(m, "__version__", "?"))
        except Exception as ex:
            check(f"module {mod}", False, str(ex))
    check("template.pdf present", TEMPLATE_PATH.exists(),
          "copy your blank form here" if not TEMPLATE_PATH.exists() else "")
    check("config.json present", CONFIG_PATH.exists())
    _miss = _missing_profile_fields(cfg)
    check("profile set up", not _miss, ("run ./timesheet setup -- missing " + ", ".join(_miss)) if _miss else "")
    print(f"  [i  ] graph login cached: {'yes' if TOKEN_CACHE.exists() else 'no (only needed for the graph method)'}")
    try:
        requests.get("https://graph.microsoft.com/v1.0/$metadata", timeout=8)
        check("can reach Microsoft Graph", True)
    except Exception as ex:
        check("can reach Microsoft Graph", False, str(ex))
    print("\nAll good." if ok else "\nSome checks failed -- see above.")


# ----------------------------- argparse -------------------------------------
def main():
    p = argparse.ArgumentParser(description="Weekly Time & Attendance timesheet tool")
    sub = p.add_subparsers(dest="cmd", required=True)

    def with_week_flags(sp):
        sp.add_argument("--week", help="target week: 'next', '+2' (weeks ahead), 'last', "
                                       "or any date in it (YYYY-MM-DD); default = this week")
        sp.add_argument("--holiday", action="append", default=[], metavar="DAY",
                        help="holiday: mon/.../fri, day number (19), YYYY-MM-DD, or a range "
                             "YYYY-MM-DD..YYYY-MM-DD (repeatable)")
        sp.add_argument("--pto", action="append", default=[], metavar="DAY[:opts]",
                        help="PTO day or range with opts ptos/ptou/paid/unpaid, e.g. "
                             "wed:ptou:unpaid or 2026-08-03..2026-08-15:unpaid (repeatable)")
        sp.add_argument("--via", choices=["outlook", "graph"],
                        help="override the send method for this run")

    sp = sub.add_parser("schedule", help="schedule the email for later (server-side)")
    with_week_flags(sp)
    sp.add_argument("--at", help="send time, e.g. 'fri 2pm', '9am', '2026-07-17 14:00'; default Fri 2:00 PM")
    sp.add_argument("--through", metavar="YYYY-MM-DD",
                    help="also schedule every week from the target week through this date (e.g. a PTO stretch)")
    sp.add_argument("--weeks", type=int, help="schedule this many consecutive weeks from the target week")
    sp.set_defaults(func=cmd_schedule)

    sp = sub.add_parser("send", help="build + send right now"); with_week_flags(sp)
    sp.set_defaults(func=cmd_send)

    sp = sub.add_parser("draft", help="build + save a ready-to-send draft in Outlook (syncs to phone)")
    with_week_flags(sp)
    sp.set_defaults(func=cmd_draft)

    sub.add_parser("setup", help="choose this machine's send method (Outlook/Graph)").set_defaults(func=cmd_setup)

    sp = sub.add_parser("eml", help="write the email as a .eml file (no account/admin needed)")
    with_week_flags(sp)
    sp.add_argument("--no-open", action="store_true", help="just write the file, don't open it")
    sp.set_defaults(func=cmd_eml)

    sp = sub.add_parser("preview", help="build + print the email, send nothing"); with_week_flags(sp)
    sp.set_defaults(func=cmd_preview)

    sp = sub.add_parser("build", help="build the PDF only"); with_week_flags(sp)
    sp.set_defaults(func=cmd_build)

    sp = sub.add_parser("verify", help="send a TEST email to yourself only (proves it works)")
    sp.add_argument("--via", choices=["outlook", "graph"], help="which method to test")
    sp.add_argument("--in", dest="in_minutes", type=int, default=2, help="[graph] minutes out to schedule the test (default 2)")
    sp.add_argument("--to", help="[graph] send the test to this address instead of yourself")
    sp.add_argument("--no-wait", action="store_true", help="[graph] submit and exit, don't watch the inbox")
    sp.set_defaults(func=cmd_verify)

    sp = sub.add_parser("login", help="[graph] sign in via browser (approve MFA)")
    sp.add_argument("--device", action="store_true", help="use device-code flow instead of a browser (may be blocked by your admin)")
    sp.add_argument("--full", action="store_true", help="request the broader scope needed for server-side scheduling")
    sp.set_defaults(func=cmd_login)
    sub.add_parser("whoami", help="show signed-in account").set_defaults(func=cmd_whoami)
    sub.add_parser("logout", help="clear cached login").set_defaults(func=cmd_logout)
    sub.add_parser("doctor", help="check prerequisites").set_defaults(func=cmd_doctor)

    args = p.parse_args()
    cfg = load_config()
    try:
        args.func(args, cfg)
    except Exception as ex:
        print(f"Error: {ex}", file=sys.stderr)
        sys.exit(1)


if __name__ == "__main__":
    main()
