---
name: urgent-email-watch
description: Operate the urgent-email-watch tool (version 2) — a real-time multi-account watchdog that polls one or more mailboxes, matches new mail against per-account triggers, and fires alerts on Pushover, Windows toast, and Alexa channels. It supports two providers: Microsoft Graph (delegated device-code OAuth2, for both Microsoft 365 work mail and personal Outlook.com / Hotmail) and Gmail (IMAP + SMTP with an App Password). Use this skill when the user wants to install, configure, set up, test, or troubleshoot urgent-email-watch; add, edit, or retire a mailbox account or an email trigger; understand the channels model; check the alert log or pending alerts; or work out why an alert did or did not fire.
---

# urgent-email-watch (version 2)

## What this tool is

`urgent-email-watch` is a self-contained Windows tool that watches one or more
mailboxes and raises an alert the moment an important email arrives. Each poll
cycle it loops over every configured account, polls that mailbox, matches each
new message against that account's hand-editable triggers, and fires
notifications on up to three channels: Pushover (phone push), Windows toast
(desktop popup), and Alexa (spoken notification on Echo devices via the
"Notify Me" skill).

It is a single Python script (`watchdog.py`, standard library only) plus a
PowerShell toast renderer (`Show-Toast.ps1`). It authenticates to each mail
provider itself — no external service hosts it.

## Package layout

All files sit in one folder:

- `watchdog.py` — the engine. Multi-account polling, the `MailProvider`
  abstraction, both providers, trigger matching, dispatch, and the four
  command-line modes.
- `Show-Toast.ps1` — renders the Windows desktop toast.
- `config.json` — notifications, the shared Microsoft app registration, poll
  interval, state directory, and the `accounts` list (each account embeds its
  own triggers). Created by copying `config.example.json`.
- `config.example.json` — the multi-account template, with one worked
  `microsoft` account and one worked `gmail` account.
- `Install-UrgentEmailWatch.ps1` — Windows installer / scheduled-task setup.
- `README.md` — full human setup guide.
- `LICENSE` — MIT.

State files live in the configured `state_dir` (default
`%LOCALAPPDATA%\UrgentEmailWatch\`). Most are per-account, keyed by the
account's slug: `state_<slug>.json` (poll cursor), `pending_<slug>.json`
(15-minute re-fire tracker), `token_<slug>.json` (Microsoft refresh token).
`alert_log.json` is a single shared audit log; each entry names its account.

## The multi-account model

`config.json` has an `accounts` array. Each account ("instance") is one
mailbox and carries its own `provider`, connection details, `scan_folders`,
`channels`, and `triggers`. Every poll cycle the watchdog iterates all
accounts. State is namespaced per account by a slug derived from the account
`name`, so a message id in one mailbox never collides with one in another.
Providers can be mixed freely in one config.

## Providers

Both providers implement a common `MailProvider` interface inside
`watchdog.py`: `connect()`, `list_new_messages(cursor)`,
`is_message_read(message_id)`, and `send_mail(to, subject, body)`. Every
provider returns messages in the same normalized dict shape (`id`, `subject`,
`from_address`, `from_name`, `body_preview`, `received`, `is_read`, `folder`).

- **microsoft** — Microsoft Graph with the **delegated device-code OAuth2**
  flow. One public-client app registration (config: `microsoft.client_id`,
  `authority` normally `.../common`) serves every Microsoft account, work or
  personal. No client secret, no admin consent. Each account is signed in once
  via the device-code flow (the `--setup` command), which saves a refresh
  token; the token is then refreshed silently on every run, and a rotated
  refresh token is re-saved if the authority returns one. Graph calls use the
  delegated `/me/` endpoints.
- **gmail** — IMAP read + SMTP send, authenticated with a Google **App
  Password**. New mail is tracked by IMAP **UID**: the cursor is a per-folder
  map of the highest UID seen. Sending (for the Pushover channel) goes through
  Gmail SMTP with STARTTLS. Standard library only — `imaplib`, `smtplib`,
  `email`.

## How it behaves

- **Polling.** Every `poll_interval_sec` the watchdog runs one cycle across
  all accounts. For each account it lists messages newer than that account's
  cursor in each `scan_folders` entry, normalizes them, and matches them
  against that account's enabled triggers.
- **Matching.** A trigger fires if an email matches ANY of: a domain in
  `sender_domains` (suffix match on the From address), an address in
  `sender_addresses` (exact match), a keyword in `subject_keywords`
  (case-insensitive substring), or a keyword in `body_keywords` (substring of
  the body preview). The first matching trigger wins.
- **Channels.** A channel fires for a matched email only if it is enabled in
  the resolved channel set AND its global destination in `notifications` is
  configured. The resolved set is the account's `channels` block, optionally
  overridden by a `channels` block on the matching trigger. The account
  `channels` block defaults to all-on if omitted; a missing key inside any
  `channels` block also defaults to `true`.
- **Idempotency.** Every alert is keyed by (account slug, message id). A
  message already handled for its account (in that account's
  `pending_<slug>.json`, or in `alert_log.json` with the same account slug)
  never fires again.
- **Re-fire.** Roughly every 5 minutes the watchdog checks each account's
  pending alerts. If an alerted email is still unread 15+ minutes after the
  first fire, it re-fires the Pushover channel once, then stops. Reading the
  email in the mailbox acknowledges and clears the pending entry.
- **Resilience.** An exception inside one account's poll pass is logged (with
  the account slug) and the watchdog moves on to the next account; one broken
  account never kills the loop or blocks the others.

## Operating the tool

### Install

Run `Install-UrgentEmailWatch.ps1` in PowerShell. It checks for Python 3,
copies `config.example.json` to `config.json` (never overwriting an existing
one), and registers a scheduled task `UrgentEmailWatch` that runs the watchdog
at logon. It is idempotent — safe to re-run. After install, the user must edit
`config.json` and then run `--setup`.

### Command-line modes

- `python watchdog.py` — run the forever loop across all accounts (used by the
  scheduled task).
- `python watchdog.py --setup` — for every Microsoft account missing a valid
  cached refresh token, run the interactive device-code sign-in and save the
  token. Gmail accounts need no interactive setup. Run this after editing
  `config.json`, and again only if a token is later revoked or expires.
- `python watchdog.py --test` — validate config; for each account verify
  auth/connection works and list that account's triggers; exit. Run this first
  after editing `config.json`.
- `python watchdog.py --once` — run a single poll pass across all accounts and
  exit.
- `python watchdog.py --help` — print usage.

### Add a mailbox account

Add an object to the `accounts` array in `config.json`: a unique `name`, a
`provider` (`microsoft` or `gmail`), the provider's connection fields
(`user_email` for microsoft; `email` + `app_password` for gmail),
`scan_folders`, optionally `channels`, and a `triggers` list. For a `microsoft`
account, run `python watchdog.py --setup` afterward to sign it in. The watchdog
reads `config.json` only at startup, so restart it (or its scheduled task)
after editing `accounts`.

### Add or retire a trigger

Edit the relevant account's `triggers` array in `config.json`. Add an object
with `name`, `enabled`, the four match-list fields, and optionally a
`channels` block (which overrides the account default for that trigger). To
retire a trigger, set `"enabled": false` and leave the object as a record.
Restart the watchdog to pick up changes.

### Check what happened

- `alert_log.json` — did an alert fire for a given message id, in which
  account, and what did each channel return? (Shared across accounts; each
  entry has an `account` / `account_slug`.)
- `pending_<slug>.json` — which of that account's alerts are awaiting
  acknowledgment or re-fire?
- `state_<slug>.json` — that account's poll cursor; if `updated_at` is stale,
  the watchdog may not be running.

## Diagnosing a missed alert

1. Identify which account should have caught it. Check `alert_log.json` for
   the email's message id under that account's slug. If present, the watchdog
   detected and dispatched it — any failure is downstream of dispatch.
2. Confirm the trigger that should have matched is `enabled: true` on the
   right account and its match rules actually cover the email's From address /
   subject.
3. Confirm the relevant channel is enabled in the resolved channel set (the
   account's `channels`, or the trigger's `channels` if it has one) AND its
   destination in `notifications` is configured.
4. Check `state_<slug>.json` `updated_at` — if stale, the watchdog is not
   running; start the scheduled task.
5. Run `python watchdog.py --test` to confirm config is valid and every
   account's auth/connection works. For a failing Microsoft account, run
   `python watchdog.py --setup`.

## Configuration reference

`config.json` top-level keys: `notifications` (`pushover.gateway_email`,
`notify_me.access_code`, `toast.enabled`), `microsoft` (`client_id`,
`authority`), `poll_interval_sec`, `state_dir`, and `accounts`. Each account:
`name`, `provider`, connection fields, `scan_folders`, `channels`, `triggers`.
Each trigger: `name`, `enabled`, `sender_domains`, `sender_addresses`,
`subject_keywords`, `body_keywords`, optional `channels`, optional
`rationale`. The `_comment` block in `config.example.json` documents every
field. See `README.md` for the full setup walkthrough including the Microsoft
Entra app registration, Gmail App Passwords, Pushover, and Alexa "Notify Me".
