# Time & Attendance timesheet tool

Fills the weekly **Daily Attendance Record** PDF (a real fillable form) for a
Sunday→Saturday payroll week and gets it to the payroll team — by **sending now**,
**staging a ready-to-send draft/scheduled email**, or **exporting a `.eml` file**.
Dates, holidays, PTO and the weekly hour total are filled exactly and recomputed
automatically, so each week is one short command.

**Built for anyone on the team to reuse.** Nothing personal is committed — your name,
Life #, email, recipients, login token and your blank form PDF all stay **local and
git-ignored**. On first run the tool asks for your details and saves them to
`profile.json` on your machine only. The email it produces is:

- **To:** your payroll address (default `GGSTimeAttendance@mssm.edu`)  **Cc:** your supervisor
- **Subject:** Time & Attendance Sheet submission
- **Body:** the fixed template with the payroll period (mm/dd to mm/dd) and your name filled in.

---

## Three ways to send (pick one per machine)

| Method | Needs | Friday 2 PM scheduling | Notes |
|---|---|---|---|
| **Outlook** (recommended) | Outlook app signed into your mssm account | Yes — via Outlook's **Schedule Send** (New) or a saved draft you send | No admin/IT approval. Sends through your own Outlook. |
| **Graph** (API) | Org/IT to approve the app's mail permission | Yes — true **server-side** auto-send (Mac can be off) | At Mount Sinai this needs IT admin consent. |
| **`.eml` file** | Nothing | No (you press Send) | Universal fallback; works in any state. |

The tool remembers your choice per machine. Override any run with `--via outlook|graph`.

---

## 1. Install (per machine — Apple Silicon or Intel)

```bash
git clone https://github.com/iamakhilverma/ismms-timesheet-tool.git timesheet
cd timesheet
./setup.sh           # isolated .venv + deps (does NOT touch system Python)
./timesheet setup    # enter your name, Life #, email, recipients; pick Outlook/Graph

# Add YOUR blank fillable "Daily Attendance Record" PDF (the standard form from your
# department) as template.pdf -- it stays local (git-ignored). If yours is signed, your
# signature carries onto every generated sheet.
cp "/path/to/your blank Daily Attendance Record.pdf" template.pdf

./timesheet doctor   # sanity-check
./timesheet verify   # send a test to yourself only
```

`setup.sh` needs Python 3.9+ (ships with Xcode Command Line Tools — if missing it tells
you to run `xcode-select --install`). Everything installs into `./.venv` only.

### Your private files (never committed)

These live only on your machine — all git-ignored, so the repo is safe to share/publish:

| File | What it holds | Created by |
|---|---|---|
| `profile.json` | your name, Life #, email, recipients, your holidays/PTO | `timesheet setup` |
| `template.pdf` | your blank Daily Attendance Record form (+ signature) | you copy it in |
| `.token_cache.json` | your cached Microsoft login (Graph method) | `timesheet login` |
| `.backend` | this machine's method choice (outlook/graph) | `timesheet setup` |

`config.json` (committed) holds only **non-personal** defaults; `profile.json` overrides
it with your info. Copy `profile.example.json` → `profile.json` to edit by hand instead
of running `setup`.

### Run it from anywhere

You **don't** have to be in this folder — the tool reads its own `config.json` /
`template.pdf` by location, not your current directory. `setup.sh` offers to add a global
**`timesheet`** alias to `~/.zshrc`; accept it, then `source ~/.zshrc` (or open a new
terminal) and run it from anywhere:

```bash
timesheet draft --week next
```

If you skipped it, add it yourself: `alias timesheet="/full/path/to/timesheet"` in
`~/.zshrc`. (Examples below show `./timesheet`; with the alias, just drop the `./`.)

---

## 2. Sign-in / login cases

### Outlook method (recommended, no admin)
1. Open the **Outlook app** and sign into **your work account** (the email you entered in
   `setup`) — your normal mailbox login + MFA; you're adding your own inbox, nothing for
   IT to approve. Let it sync.
2. First time the tool controls Outlook, macOS asks for **Automation permission** —
   click **OK** (or System Settings → Privacy & Security → Automation → allow it).
3. Test it (sends only to you, kept in your inbox): `./timesheet verify --via outlook`
4. *(Optional, cosmetic)* Outlook appends a default **"Get Outlook for Mac"** signature
   to new mail. To remove it: **Outlook → Settings → Signatures** → delete that signature,
   or set **New messages → (none)**. The tool fills the body; Outlook adds the signature
   afterward, so this is an Outlook setting (not something the tool controls).

**New Outlook vs Classic ("Legacy") Outlook** — for everyday use you don't need to change
anything. The **only** reason to use Classic is **bulk-drafting many weeks at once**.

| | Send / single-week `draft` | Bulk `draft --through` / `--weeks` |
|---|---|---|
| **New Outlook** (default) | ✅ opens a filled compose window → Schedule Send or save to Drafts | ❌ would open one window per email — the tool refuses for >1 week |
| **Classic / Legacy** | ✅ saves the draft straight to Drafts (syncs to phone) | ✅ saves all drafts silently to your Drafts folder |

**Bottom line:** stay in **New Outlook** for normal weekly use. Switch to **Classic** only
to **bulk-draft a range** — toggle **"New Outlook" off** (top-right of Outlook; it
restarts), run the bulk command, then switch back if you like. The bulk drafts sync
everywhere, so to schedule them for Friday 2 PM, open them in **New Outlook / OWA / your
phone** (which have *Schedule Send*; Classic's UI doesn't).

> Switching between New and Classic uses the **same account and mailbox** — it's just the
> UI/automation layer. Your mail, drafts, and settings are unchanged.

### Graph method (only if your org approves it)
```bash
./timesheet login          # opens your browser (auth-code flow), approve MFA
./timesheet login --full   # broader scope needed for server-side scheduling
```
- The **browser** flow is used (it passes Mount Sinai's Conditional Access; the older
  device-code flow is blocked — `--device` exists only as a fallback).
- Your org may require **admin approval** of the app ("Microsoft Graph Command Line
  Tools"). If sign-in says a request was sent to your admin, you must wait for approval.
- Login is cached (`.token_cache.json`, git-ignored); you re-approve MFA only rarely.

---

## 3. The weekly protocol (how you actually use it)

Run **one** command. By default it targets **this** week (the Sun–Sat week containing
today). To generate the submission for a **different** week — next week, somewhere in the
future, or a past week you missed — add `--week`:

| You want… | Add | Example |
|---|---|---|
| This week (default) | *(nothing)* | `./timesheet draft` |
| **Next** week | `--week next` | `./timesheet draft --week next` |
| **N weeks ahead** | `--week +N` | `./timesheet draft --week +3` |
| **Last** / a past week | `--week last` or `--week -N` | `./timesheet draft --week last` |
| A **specific** week | `--week YYYY-MM-DD` (any day in it) | `./timesheet draft --week 2026-07-12` |

`--week` works on **every** command below (`draft`, `schedule`, `send`, `preview`,
`build`) and combines with `--holiday` / `--pto`, e.g. `./timesheet draft --week next
--holiday fri`. Tip: run `./timesheet preview --week next` first to see exactly what that
week's sheet and email will contain before staging it.

### A. Stage a draft / scheduled email — **recommended**
```bash
./timesheet draft               # this week
./timesheet draft --week next    # next week's submission
```
- **New Outlook:** opens the filled compose window → click the **▾ next to Send →
  Schedule Send → Friday 2:00 PM** (Exchange sends it then, even if your Mac is off),
  **or** save to Drafts (⌘S) and tap Send from your **phone** Friday.
- **Classic Outlook:** saves it to **Drafts** (syncs to your phone) → tap Send Friday.

### B. True server-side schedule (Graph method only)
```bash
./timesheet schedule                 # delivers Friday 2:00 PM automatically
./timesheet schedule --at "fri 9am"  # different time
```

### C. Send right now
```bash
./timesheet send
```

### D. Universal file (any method/state)
```bash
./timesheet eml      # writes + opens a .eml; open in Outlook (desktop/web) and press Send
```

### Preview without sending anything
```bash
./timesheet preview      # prints the sheet summary + the email
./timesheet build        # just writes the filled PDF to output/
```

---

## 4. Use cases (holidays, PTO, ranges, multiple weeks)

Pick the day with a **weekday** (`mon`…`fri`), a **day number** (`19`), a **date**
(`2026-06-19`), or a **range** (`2026-08-03..2026-08-15`). Weekends never carry hours.

```bash
# A holiday this week (code H, blank in/out, 7.5h, total stays 37.5)
./timesheet draft --holiday fri
./timesheet draft --holiday 19

# PTO — scheduled & paid by default (PTOS, 7.5h)
./timesheet draft --pto wed

# PTO unpaid -> that day has no hours, weekly total drops
./timesheet draft --pto wed:unpaid

# PTO unscheduled / unpaid
./timesheet draft --pto wed:ptou:unpaid

# Combine, repeatable
./timesheet draft --holiday fri --pto mon

# A stretch of PTO across several weeks, staged all at once:
./timesheet schedule --through 2026-08-15 --pto 2026-08-03..2026-08-15:unpaid
./timesheet schedule --weeks 4            # next 4 weeks in one go

# Target a different week (see the --week table in §3) and combine with time off:
./timesheet draft --week next --holiday fri
./timesheet draft --week 2026-07-12 --pto wed
./timesheet schedule --at "2026-07-17 13:30"
```

**Absence codes** (per the form): `H` = Holiday, `PTOS` = PTO Scheduled,
`PTOU` = PTO Unscheduled. Paid days keep 7.5h; unpaid days are left blank and lower the
total.

---

## 5. Set holidays / PTO once (so you don't pass flags)

Edit `config.json` and the tool applies them automatically every run. Both are empty by
default (works for any year). `config.json` ships with `_holidays_example` /
`_pto_example` showing the exact format — copy into the real `holidays` / `pto` keys.

```jsonc
// holidays: ISO dates or A..B ranges (weekends in a range are skipped)
"holidays": ["2026-06-19", "2026-12-24..2026-12-25"],

// pto: each entry is {date} OR {from,to}, with code PTOS/PTOU and paid true/false
"pto": [
  { "from": "2026-08-03", "to": "2026-08-15", "code": "PTOU", "paid": false },
  { "date": "2026-07-10", "code": "PTOS", "paid": true }
]
```

You can put holidays/PTO in `profile.json` instead (also auto-applied) if you'd rather
keep them with your personal info. `config.json` holds shared defaults (subject/body
template, work hours/times, holiday code, default send time, and the org's holiday list);
your name, Life #, email and recipients live in `profile.json`.

**Auto-import holidays from the official PDF:** instead of typing dates, point the tool at
your department's holiday-schedule PDF — it reads the *observed* dates:

```bash
timesheet holidays-from-pdf "Holiday_Schedule_2026-MSH.pdf"                 # preview
timesheet holidays-from-pdf "Holiday_Schedule_2026-MSH.pdf" --write config   # save
```

**Bulk-draft a whole stretch** (e.g. the rest of the year) in one command. **This is the
one feature that needs Classic Outlook** — switch the "New Outlook" toggle **off** first
(New Outlook would open a window per email; Classic saves them silently to Drafts, synced
to your phone). Switch back to New afterward to Schedule-Send them. See
[New vs Classic](#2-sign-in--login-cases) above.

```bash
timesheet draft --week 2026-06-28 --through 2026-12-31   # every week in the range
timesheet draft --weeks 6                                 # the next 6 weeks
```

The drafts sync to your phone / OWA / New Outlook, where you open each and **Schedule
Send → that week's Friday 2 PM**.

---

## 6. Command reference

| Command | What it does |
|---|---|
| `setup` | enter your profile (name/Life#/email/recipients) + choose method (outlook/graph) |
| `draft` | stage the email in Outlook (compose window / saved draft); **bulk** a range with `--through`/`--weeks` |
| `holidays-from-pdf` | extract holiday dates from a holiday-schedule PDF (`--write profile`/`config`) |
| `schedule` | Graph: server-side Friday 2 PM. Outlook: stage draft(s). Supports `--through` / `--weeks` |
| `send` | send right now |
| `eml` | write the email as a `.eml` file and open it |
| `preview` | print the sheet summary + email, send nothing |
| `build` | write just the filled PDF |
| `verify` | send a test **to you only** (proves it works; nothing to payroll) |
| `login` / `whoami` / `logout` | Graph sign-in / show account / clear cached login |
| `doctor` | check prerequisites |

Common flags: `--week`, `--holiday`, `--pto`, `--at`, `--through`, `--weeks`, `--via`.

---

## 7. Troubleshooting

- **"Outlook has no mail account signed in"** (Classic Outlook) → sign into Outlook, let
  it sync, retry.
- **New Outlook shows 0 accounts to automation** → that's expected; sending/compose still
  work. The tool does not block on it.
- **macOS blocked controlling Outlook** → System Settings → Privacy & Security →
  Automation → allow your terminal/Python to control Microsoft Outlook.
- **Graph sign-in says a request was sent to your admin** → the org requires admin
  approval of the app; wait for IT, or use the Outlook method instead.
- **Verify can't auto-confirm** in New Outlook → it can't read Sent Items via automation;
  just check your inbox.
- **"Get Outlook for Mac" appears at the bottom of the email** → that's Outlook's default
  signature, not the tool. Remove it in Outlook → Settings → Signatures (see step 4 above).
- **"Your profile isn't set up yet"** → run `./timesheet setup` (or copy
  `profile.example.json` → `profile.json` and fill it in).
- **"No template.pdf found"** → copy your blank Daily Attendance Record form to
  `template.pdf` in this folder.

---

## Privacy

**No personal data is committed.** Your name, Life #, email, recipients, Microsoft login
token, per-machine choice, and every PDF (your blank form and all generated sheets) are
git-ignored — `profile.json`, `template.pdf`, `*.pdf`, `.token_cache.json`, `.backend`.
`config.json` carries only shared, non-personal defaults. So the repo is **safe to share
with colleagues or make public**, and each person's data stays on their own machine.

> **Making an existing repo public?** Git *history* can still contain files you removed
> later. If personal data was ever committed (e.g. an early `template.pdf` or a
> `config.json` with your name), scrub history first — the simplest is to delete the
> GitHub repo and re-create it from the current clean state, or rewrite history with
> `git filter-repo`. (For a brand-new clone of this clean repo, there's nothing to scrub.)
