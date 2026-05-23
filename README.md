# urgent-email-watch (version 2)

A real-time urgent-email watchdog for Windows. **Version 2 watches many
mailboxes at once** and supports two mail providers.

## What this tool does

`urgent-email-watch` polls one or more mailboxes every 60 seconds
(configurable). Each new message is matched against a list of **triggers** you
define — sender domains, sender addresses, subject keywords, and body
keywords. When a trigger matches, the tool fires an alert on up to three
independent channels:

- **Pushover** — sends an email to your Pushover email-to-push gateway, which
  pushes a notification to the Pushover app on your phone.
- **Windows toast** — a desktop notification on the machine running the tool.
- **Alexa** — posts to the "Notify Me" Alexa skill, which queues a spoken
  notification on your Echo devices.

It is idempotent (an alert fires once per message), resilient (a transient
error in one account never kills the watcher or blocks the other accounts),
and keeps an audit log of every alert. If an alerted email is still unread
after 15 minutes, it re-fires the Pushover channel once.

The tool is self-contained: a single Python script plus a PowerShell toast
renderer. No databases, no web server, no pip packages.

### The multi-account model

The config file has an `accounts` list. Each account ("instance") is one
mailbox and carries its own:

- **provider** — `microsoft` (Microsoft 365 work mail *and* personal
  Outlook.com / Hotmail) or `gmail`.
- **connection details** — how to reach that mailbox.
- **triggers** — the emails *that account* cares about.
- **channels** — which of Pushover / toast / Alexa *that account* may fire.

Every poll cycle the watchdog loops over **all** accounts: for each one it
polls that mailbox, matches that account's triggers, and fires that account's
enabled channels. State (the poll cursor, the message-id idempotency record,
the 15-minute pending tracker, and Microsoft refresh tokens) is namespaced
**per account**, so a message id in one mailbox can never collide with one in
another. You can mix providers freely — for example one Microsoft work
mailbox, one personal Outlook.com mailbox, and one Gmail account, all in one
config.

---

## Supporting software

### Python 3

The watchdog is a Python 3 script. Install Python with winget:

```
winget install Python.Python.3.12
```

Or download the installer from <https://www.python.org/downloads/windows/>.
During installation, tick **"Add python.exe to PATH"**.

**No pip packages are needed.** The watchdog uses only the Python standard
library (`json`, `urllib`, `imaplib`, `smtplib`, `email`, `subprocess`, etc.).

### PowerShell

PowerShell is **built into Windows** — nothing to install. It is used to render
the desktop toast and to run the installer. The toast renderer uses only
built-in Windows APIs; no third-party modules (such as BurntToast) are
required.

---

## Microsoft accounts (Microsoft 365 and Outlook.com)

Version 2 authenticates to Microsoft Graph with the **delegated device-code
OAuth2 flow**. You register **one** application — a *public client* — and that
single registration works for **every** Microsoft account you add, whether it
is a Microsoft 365 / Entra work account or a personal Outlook.com / Hotmail
account. There is **no client secret** and **no admin consent** to deal with.

### Register the app once

In the Azure portal (<https://portal.azure.com>):

1. Go to **Microsoft Entra ID** → **App registrations** → **New
   registration**.
2. Name it (e.g. `urgent-email-watch`).
3. Under **Supported account types**, choose
   **"Accounts in any organizational directory (Any Microsoft Entra ID
   tenant - Multitenant) and personal Microsoft accounts (e.g. Skype,
   Xbox)"**. This is what lets the same app sign in both work and personal
   accounts.
4. Leave the redirect URI blank. Click **Register**.
5. On the app's **Overview** page, copy the **Application (client) ID**.
6. Go to **Authentication**. Scroll to **Advanced settings** → **Allow public
   client flows** and set it to **Yes**. Save. (The device-code flow is a
   public-client flow; this switch is required.)
7. Go to **API permissions** → **Add a permission** → **Microsoft Graph** →
   **Delegated permissions** (not Application). Add:
   - `Mail.Read`
   - `Mail.Send`
   You do **not** need to click "Grant admin consent" — delegated
   `Mail.Read` / `Mail.Send` are user-consentable, and each user consents to
   them when they sign in during `--setup`.
8. You do **not** need a client secret. Device-code auth uses none.

### Put the client ID in config.json

```json
"microsoft": {
  "client_id": "<the Application (client) ID you copied>",
  "authority": "https://login.microsoftonline.com/common"
}
```

`authority` should stay as `https://login.microsoftonline.com/common` —
`/common` is what allows both work and personal accounts.

### Add each Microsoft mailbox as an account

Each Microsoft account entry only needs `user_email` and its `scan_folders` —
the `client_id` is shared from the `microsoft` block:

```json
{
  "name": "Work Mailbox",
  "provider": "microsoft",
  "user_email": "you@yourcompany.com",
  "scan_folders": ["inbox"],
  "channels": { "pushover": true, "toast": true, "alexa": true },
  "triggers": [ ... ]
}
```

### Sign in once per account

After filling in `config.json`, run:

```
python watchdog.py --setup
```

For each Microsoft account, this prints a short message: open a URL
(`https://microsoft.com/devicelogin`) in any browser, type the displayed code,
and sign in with **that account's** credentials. The watchdog then saves a
refresh token for that account (`token_<slug>.json` in the state directory)
and uses it silently on every run afterward. You only do this once per
account; re-run `--setup` only if a token is later revoked or expires.

---

## Gmail accounts

Gmail accounts use **IMAP** to read and **SMTP** to send, authenticated with a
**Google App Password**. No app registration, no OAuth, no interactive setup.

1. Turn on **2-Step Verification** for the Google account
   (<https://myaccount.google.com/security>). App Passwords are only available
   once 2-Step Verification is on.
2. Go to **App Passwords** (<https://myaccount.google.com/apppasswords>),
   create a new app password (name it `urgent-email-watch`), and copy the
   16-character password Google shows you.
3. Put the address and the App Password in the account entry:

   ```json
   {
     "name": "Personal Gmail",
     "provider": "gmail",
     "email": "you@gmail.com",
     "app_password": "xxxx xxxx xxxx xxxx",
     "scan_folders": ["inbox"],
     "channels": { "pushover": true, "toast": false, "alexa": true },
     "triggers": [ ... ]
   }
   ```

`scan_folders` for Gmail can name the Inbox (`inbox`) or any Gmail **label**
(labels are addressable as IMAP folders — e.g. `"Important"`). Gmail accounts
are skipped by `--setup` because they need no interactive sign-in.

---

## Pushover setup

Pushover delivers push notifications to your phone, independent of any cellular
carrier.

1. Create a Pushover account at <https://pushover.net> and install the Pushover
   app on your phone.
2. On your Pushover account page, find the **email-to-push gateway address**.
   It is listed under your devices — Pushover gives you an address like
   `something@pomail.net` that turns inbound email into a push notification.
3. Put that address in `config.json` under `notifications`:

   ```json
   "notifications": {
     "pushover": { "gateway_email": "something@pomail.net" }
   }
   ```

When an alert fires, the watchdog sends an email to that gateway **using the
matched account's own provider** (Graph `sendMail` for Microsoft accounts,
Gmail SMTP for Gmail accounts). **The email subject becomes the push
notification title; the email body becomes the message.** Leave
`gateway_email` blank to disable the Pushover channel for every account.

---

## Alexa "Notify Me" setup

The Alexa channel uses the free "Notify Me" skill to put spoken notifications on
your Echo devices.

1. Open the **Alexa app** on your phone. Go to **More** → **Skills & Games**.
2. Search for **"Notify Me"** and **enable** the skill.
3. The skill **emails you an access code** — a string that looks like
   `nmac.XXXXXXXXXXXX`.
4. Paste that access code into `config.json` under `notifications`:

   ```json
   "notifications": {
     "notify_me": { "access_code": "nmac.XXXXXXXXXXXX" }
   }
   ```

When an alert fires, every Echo device on your Alexa account flashes a **yellow
notification ring**. To hear the queued notifications, say
**"Alexa, what are my notifications?"** and Alexa reads them aloud.

It is a **free** service with a **250-character limit** per message. The
watchdog caps each message at 245 characters. Leave `access_code` blank to
disable the Alexa channel for every account.

---

## Configuring accounts and triggers

Everything lives in `config.json` next to the script. There is no separate
triggers file in v2 — each account embeds its own `triggers` list.

### Config schema (overview)

```json
{
  "notifications": {
    "pushover":  { "gateway_email": "something@pomail.net" },
    "notify_me": { "access_code": "nmac.XXXXXXXX" },
    "toast":     { "enabled": true }
  },
  "microsoft": {
    "client_id": "<Entra app client id>",
    "authority": "https://login.microsoftonline.com/common"
  },
  "poll_interval_sec": 60,
  "state_dir": "%LOCALAPPDATA%\\UrgentEmailWatch",
  "accounts": [ { ... }, { ... } ]
}
```

### An account entry

```json
{
  "name": "Work Mailbox",
  "provider": "microsoft",
  "user_email": "you@yourcompany.com",
  "scan_folders": ["inbox", "Important"],
  "channels": { "pushover": true, "toast": true, "alexa": true },
  "triggers": [ ... ]
}
```

- `name` — any label. It is also slugified to name that account's state files
  (`state_<slug>.json`, `pending_<slug>.json`, `token_<slug>.json`).
- `provider` — `microsoft` or `gmail`.
- Connection fields — `user_email` for `microsoft`; `email` and
  `app_password` for `gmail`.
- `scan_folders` — mail folders/labels to scan. `inbox` is the well-known
  Inbox; other names are matched by display name (Microsoft) or label (Gmail).
- `channels` — which channels this account may fire.
- `triggers` — this account's trigger list.

### How channels default to all-on

The `channels` block on an account is **optional**. If you omit it entirely,
the account defaults to **all three channels on** (`pushover`, `toast`,
`alexa` all `true`). A key missing *inside* the block also defaults to `true`.

A channel actually fires for a matched email only if **both** of these are
true:

1. the channel is enabled in the account's (or trigger's) `channels`, **and**
2. the matching global destination in `notifications` is configured —
   `pushover.gateway_email` is set, `notify_me.access_code` is set, or
   `toast.enabled` is `true`.

So `notifications` is the master "is this channel even possible" switch, and
each account/trigger `channels` block is the per-mailbox selection.

### A trigger entry

```json
{
  "name": "Bank alert",
  "enabled": true,
  "sender_domains": ["yourbank.example.com"],
  "sender_addresses": ["alerts@yourbank.example.com"],
  "subject_keywords": ["large transaction", "fraud"],
  "body_keywords": ["suspicious"],
  "channels": { "pushover": true, "toast": true, "alexa": true },
  "rationale": "Free-text note; ignored by the watchdog."
}
```

**Match logic** — a trigger fires if an incoming email matches **any** of:

- `sender_domains` — the From address ends with `@<domain>` or `.<domain>`.
- `sender_addresses` — the From address exactly equals an entry
  (case-insensitive).
- `subject_keywords` — a keyword is a case-insensitive substring of the
  subject.
- `body_keywords` — a keyword is a case-insensitive substring of the body
  preview.

**Trigger-level `channels` (optional override).** A trigger MAY include its
own `channels` block. If it does, that block overrides the account default
**for that trigger only**. If a trigger has no `channels` block, it inherits
the account's `channels`. (As always, a channel still only fires if its
`notifications` destination is configured.)

### Add an account

Add another object to the `accounts` array, give it a unique `name`, set its
`provider` and connection fields, channels, and triggers. If it is a
`microsoft` account, run `python watchdog.py --setup` afterward to sign it in.
The watchdog re-reads `config.json` only at startup, so restart the watchdog
(or its scheduled task) after editing `accounts`.

### Add or retire a trigger

Add an object to that account's `triggers` array, or set `"enabled": false`
on one to retire it without deleting it. Restart the watchdog to pick up the
change.

---

## Install and run

1. Put this folder somewhere permanent (e.g. `C:\Tools\urgent-email-watch`).
2. Run the installer in PowerShell:

   ```
   powershell -NoProfile -ExecutionPolicy Bypass -File Install-UrgentEmailWatch.ps1
   ```

   The installer checks for Python 3, copies `config.example.json` →
   `config.json` (it never overwrites an existing `config.json`), and
   registers a Windows Scheduled Task named `UrgentEmailWatch` that runs the
   watchdog **at logon**. The installer is idempotent — re-run it any time.
3. Edit `config.json` — `notifications`, the `microsoft` app `client_id`, and
   the `accounts` list.
4. **Run setup once.** This is required for Microsoft accounts:

   ```
   python watchdog.py --setup
   ```

   For each Microsoft account it prints a URL and a code; open the URL, enter
   the code, and sign in as that account. The refresh token is saved. Gmail
   accounts are skipped (they use the App Password from `config.json`).
5. Test the setup:

   ```
   python watchdog.py --test
   ```

   This validates `config.json`, verifies auth/connectivity for **every**
   account, and lists each account's triggers. If it prints "Test passed", you
   are ready.
6. Optionally run a single poll pass across all accounts:

   ```
   python watchdog.py --once
   ```

7. Start the watchdog now without waiting for the next logon:

   ```
   Start-ScheduledTask -TaskName UrgentEmailWatch
   ```

From then on the watchdog **runs continuously** — it starts automatically every
time you log on, polls every mailbox forever, and the scheduled task restarts
it if it ever stops.

### Command-line modes

| Command | What it does |
|---|---|
| `python watchdog.py` | Run the forever loop across all accounts (this is what the scheduled task uses). |
| `python watchdog.py --setup` | Sign in every Microsoft account missing a valid token (device-code flow). Run after editing config. |
| `python watchdog.py --test` | Validate config, verify each account's auth/connection, list each account's triggers, exit. |
| `python watchdog.py --once` | Run exactly one poll pass across all accounts, then exit. |
| `python watchdog.py --help` | Print usage. |

### State files

The watchdog keeps its state in the configured `state_dir` (by default
`%LOCALAPPDATA%\UrgentEmailWatch\`). Most files are **per account**, named with
the account's slug:

- `state_<slug>.json` — that account's poll cursor (a timestamp for Microsoft,
  a per-folder IMAP UID map for Gmail).
- `pending_<slug>.json` — that account's alerts awaiting acknowledgment or the
  15-minute re-fire.
- `token_<slug>.json` — that Microsoft account's saved refresh token (Gmail
  accounts have no token file).
- `alert_log.json` — the **shared** audit log of every alert fired across all
  accounts (last 500); each entry records which account it belongs to.

---

## Troubleshooting

**`--setup` does not finish for a Microsoft account.**
The sign-in code expires after about 15 minutes. Open
`https://microsoft.com/devicelogin` promptly, enter the code, and complete
sign-in. If you see `authorization_declined`, you cancelled the consent prompt
— re-run `--setup`. If the app's **Allow public client flows** is not set to
**Yes**, the device-code request fails — fix that in the Azure portal under
**Authentication**.

**`--test` fails at "Verifying auth/connection" for a Microsoft account.**
Either `--setup` was never run for that account, or its refresh token was
revoked/expired. Run `python watchdog.py --setup` again. Confirm
`microsoft.client_id` is the correct **Application (client) ID** and that the
app has **Delegated** `Mail.Read` and `Mail.Send` permissions.

**`--test` fails for a Gmail account.**
The most common cause is a bad App Password. Confirm 2-Step Verification is on
for the Google account and regenerate the App Password. Paste all 16
characters into `app_password` (spaces are fine). Confirm IMAP is enabled in
Gmail settings (**See all settings** → **Forwarding and POP/IMAP**).

**Token works but no alerts fire.**
Run `python watchdog.py --test` and confirm the account's triggers are listed.
Check the From address / subject of a test email against the trigger's match
rules (domain matches are suffix matches; keyword matches are substring
matches). Send yourself a matching test email and watch `alert_log.json`.

**No Pushover notification.**
Confirm `notifications.pushover.gateway_email` is your real Pushover gateway
address and the account's `channels.pushover` is `true`. Check
`alert_log.json` — if the `pushover` channel shows `"ok": true`, the email was
accepted and the issue is downstream (Pushover account, app). Send a plain
email to the gateway address by hand to test it.

**No Alexa notification.**
Confirm `notifications.notify_me.access_code` is the full `nmac....` string and
the account's `channels.alexa` is `true`. Remember the Echo only flashes a
ring — you must say "Alexa, what are my notifications?" to hear it.

**No toast.**
Confirm `notifications.toast.enabled` is `true` and the account's
`channels.toast` is `true`. Toasts appear on the machine running the watchdog
and only in an interactive desktop session. Test the renderer directly:
`powershell -File Show-Toast.ps1 -Title "Test" -Message "Hello"`.

**One account is broken but the others should keep working.**
They do. An exception in one account's poll pass is logged (with the account
slug) and the watchdog moves straight on to the next account; the loop never
stops. Look for `[watchdog:<slug>]` lines to see which account errored.

**The watchdog is not running.**
Check the task: `Get-ScheduledTask -TaskName UrgentEmailWatch`. Start it with
`Start-ScheduledTask -TaskName UrgentEmailWatch`. The task runs at logon, so it
is not active until you have logged in.

**An email I already read keeps appearing in `pending_<slug>.json`.**
A pending alert clears when the watchdog sees the email marked read in the
mailbox, or after it re-fires once. You can also remove the entry from
`pending_<slug>.json` by hand.

**I changed `config.json` and nothing happened.**
The watchdog reads `config.json` only at startup. Restart it (or restart the
scheduled task) after editing accounts or triggers.

---

## License

MIT — see `LICENSE`.
