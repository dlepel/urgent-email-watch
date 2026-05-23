r"""
watchdog.py -- Standalone real-time urgent-email watchdog (version 2).

Watches one or more mailboxes and raises an alert the moment an important
email arrives. Version 2 adds multi-account support and two mail providers.

The config holds an `accounts` list. Each account ("instance") carries its own
provider type, its own connection details, its own `triggers`, and its own
`channels` selection. Every poll cycle the watchdog loops over ALL accounts:
for each it polls that mailbox, matches that account's triggers, and fires
that account's enabled channels. Idempotency, cursor, and pending state are
namespaced per account so message ids never collide across accounts.

Two providers are implemented behind a common MailProvider interface:

  microsoft -- Microsoft Graph, delegated device-code OAuth2. One public-client
               app registration works for BOTH M365 work accounts and personal
               Outlook.com / Hotmail accounts. No client secret, no admin
               consent. A refresh token is acquired once per account via the
               device-code flow (the --setup command) and then refreshed
               silently on every run.

  gmail     -- IMAP read + SMTP send, App Password auth. Reads via
               imaplib.IMAP4_SSL, tracks new mail by IMAP UID, and sends the
               Pushover bridge email via smtplib. Standard library only.

On a trigger match the watchdog fires alerts on up to three independent
channels:
  1. Pushover    -- sends an outbound email to a Pushover email-to-push
                    gateway address, which delivers a push to the phone app.
  2. Windows toast -- a desktop notification rendered by Show-Toast.ps1.
  3. Alexa       -- POSTs to the "Notify Me" Alexa skill REST API, which
                    queues a spoken notification on the user's Echo devices.

This is a self-contained tool. No pip packages required; standard library
only (json, urllib, imaplib, smtplib, email, subprocess, ...).

State files (default %LOCALAPPDATA%\UrgentEmailWatch\, configurable). Most are
per-account, keyed by the account slug:
  state_<slug>.json      poll cursor for one account
  pending_<slug>.json    unacknowledged alerts for one account (15-min rule)
  token_<slug>.json      microsoft refresh token for one account (--setup)
  alert_log.json         shared audit log of every alert fired (last 500)

Configuration:
  config.json  sits next to this script -- notifications, the microsoft app
               registration, poll interval, state dir, and the accounts list.
               Each account embeds its own triggers inline.

Command-line modes:
  --setup  for every microsoft account missing a valid cached refresh token,
           run the interactive device-code sign-in and save the token
  --test   validate config; for each account verify auth/connection works and
           list that account's triggers; exit
  --once   run a single poll pass across all accounts, then exit
  (none)   run the forever loop, polling every poll_interval_sec

Design notes:
  - Idempotent: every alert is keyed by (account slug, message id); never
    fires twice for the same message unless the 15-minute re-fire rule applies.
  - Resilient: an exception inside one account's pass is logged but never kills
    the loop and never blocks the other accounts.
  - The Pushover transport reuses the provider's send_mail: the email subject
    becomes the push notification title, the body becomes the message.
"""

import email
import email.message
import email.utils
import imaplib
import json
import os
import re
import smtplib
import ssl
import sys
import threading
import time
import urllib.request
import urllib.parse
import urllib.error
import subprocess
from datetime import datetime, timezone, timedelta

# --- Paths -------------------------------------------------------------------

SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(SCRIPT_DIR, "config.json")
TOAST_SCRIPT = os.path.join(SCRIPT_DIR, "Show-Toast.ps1")

# --- Constants ---------------------------------------------------------------

GRAPH_BASE = "https://graph.microsoft.com/v1.0"
DEFAULT_AUTHORITY = "https://login.microsoftonline.com/common"
GRAPH_DELEGATED_SCOPE = "Mail.Read Mail.Send offline_access"

NOTIFY_ME_URL = "https://api.notifymyecho.com/v1/NotifyMe"
NOTIFY_ME_BODY_MAX = 245     # Service limit is 250; leave buffer for ellipsis
PUSH_BODY_MAX = 1024         # Pushover allows far more; cap for tidy messages

PENDING_CHECK_INTERVAL_SEC = 300  # Re-fire check cadence (5 min)
RE_FIRE_AFTER_SEC = 15 * 60       # Re-fire push if email unread after 15 min
RE_FIRE_MAX_COUNT = 1             # Fire once on detect, once on re-fire, stop

# On first run with no state, look back this many minutes so we don't miss
# messages that arrived just before the watchdog started.
INITIAL_LOOKBACK_MIN = 5

# Token refresh: refresh when within this many seconds of expiry.
TOKEN_REFRESH_MARGIN_SEC = 300

# Device-code polling safety cap (the authority also enforces expires_in).
DEVICE_CODE_MAX_WAIT_SEC = 900

# Lock for state file mutations -- prevents races between the main pass and
# the pending checker.
_state_lock = threading.Lock()


# --- Logging -----------------------------------------------------------------

# When set, _log / _alog also append their output to this file. The installer
# runs the watchdog under pythonw.exe (no console), so file logging is the
# only way operational output -- including auth failures -- survives.
_LOG_FILE = None

# Rotate the log file once it grows past roughly this many bytes.
_LOG_MAX_BYTES = 1024 * 1024


def _write_log_line(line):
    """Append one line (with a UTC timestamp) to _LOG_FILE if it is set.
    Logging must never crash the watchdog, so all errors are swallowed."""
    if not _LOG_FILE:
        return
    try:
        stamp = datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")
        with open(_LOG_FILE, "a", encoding="utf-8") as f:
            f.write(f"{stamp} {line}\n")
    except Exception:
        pass


def init_log_file(path):
    """Point file logging at `path`, performing one-time size rotation: if the
    file already exists and exceeds ~1 MB, rename it to '<path>.1' (replacing
    any existing '.1'). All errors are swallowed -- logging is best effort."""
    global _LOG_FILE
    try:
        if os.path.exists(path) and os.path.getsize(path) > _LOG_MAX_BYTES:
            rotated = path + ".1"
            try:
                os.replace(path, rotated)
            except Exception:
                pass
    except Exception:
        pass
    _LOG_FILE = path


def _log(msg):
    """Print with a [watchdog] prefix so log lines are filterable."""
    line = f"[watchdog] {msg}"
    print(line, flush=True)
    _write_log_line(line)


def _alog(slug, msg):
    """Log a message tagged with the account slug it belongs to."""
    line = f"[watchdog:{slug}] {msg}"
    print(line, flush=True)
    _write_log_line(line)


# --- JSON / time helpers -----------------------------------------------------

def _now_iso():
    """Current time as ISO 8601 UTC with a trailing Z. Graph speaks UTC."""
    return datetime.now(timezone.utc).isoformat().replace("+00:00", "Z")


def _load_json(path, default):
    """Load a JSON file. Return default on any error."""
    try:
        with open(path, "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return default


def _save_json(path, data):
    """Atomic write: write to a .tmp file, then replace the original."""
    tmp = path + ".tmp"
    with open(tmp, "w", encoding="utf-8") as f:
        json.dump(data, f, indent=2)
    os.replace(tmp, path)


def _slugify(name):
    """Turn an account name into a filesystem-safe slug used to namespace
    that account's state files and tokens."""
    s = re.sub(r"[^a-z0-9]+", "-", str(name).strip().lower())
    s = s.strip("-")
    return s or "account"


# --- Configuration -----------------------------------------------------------

class ConfigError(Exception):
    """Raised when config.json is missing or invalid."""


def load_config():
    """Load and validate config.json. Raises ConfigError on a fatal problem.

    Returns the parsed config dict with: defaults filled in for optional
    fields, state_dir resolved to an absolute created directory, and every
    account given a unique `_slug`.
    """
    if not os.path.exists(CONFIG_FILE):
        raise ConfigError(
            f"config.json not found next to the script ({CONFIG_FILE}).\n"
            "Copy config.example.json to config.json and fill it in."
        )

    cfg = _load_json(CONFIG_FILE, None)
    if not isinstance(cfg, dict):
        raise ConfigError("config.json is not valid JSON or is not an object.")

    # --- notifications: the shared channel destinations ----------------------
    notif = cfg.get("notifications")
    if not isinstance(notif, dict):
        notif = {}
    notif.setdefault("pushover", {})
    notif["pushover"].setdefault("gateway_email", "")
    notif.setdefault("notify_me", {})
    notif["notify_me"].setdefault("access_code", "")
    notif.setdefault("toast", {})
    notif["toast"].setdefault("enabled", True)
    cfg["notifications"] = notif

    # --- microsoft: the shared public-client app registration ----------------
    ms = cfg.get("microsoft")
    if not isinstance(ms, dict):
        ms = {}
    ms.setdefault("client_id", "")
    ms.setdefault("authority", DEFAULT_AUTHORITY)
    cfg["microsoft"] = ms

    cfg.setdefault("poll_interval_sec", 60)

    # --- accounts ------------------------------------------------------------
    accounts = cfg.get("accounts")
    if not isinstance(accounts, list) or not accounts:
        raise ConfigError(
            "config.json must have a non-empty 'accounts' array. "
            "Copy config.example.json for a worked template."
        )

    seen_slugs = {}
    for idx, acct in enumerate(accounts):
        if not isinstance(acct, dict):
            raise ConfigError(f"account #{idx + 1} is not a JSON object.")

        name = str(acct.get("name") or "").strip()
        if not name:
            raise ConfigError(f"account #{idx + 1} is missing 'name'.")

        provider = str(acct.get("provider") or "").strip().lower()
        if provider not in ("microsoft", "gmail"):
            raise ConfigError(
                f"account '{name}' has unknown provider '{provider}'. "
                "Valid providers: microsoft, gmail."
            )
        acct["provider"] = provider

        # Unique slug -- disambiguate collisions with a numeric suffix.
        base_slug = _slugify(name)
        slug = base_slug
        n = 2
        while slug in seen_slugs:
            slug = f"{base_slug}-{n}"
            n += 1
        seen_slugs[slug] = True
        acct["_slug"] = slug

        # Scan folders default.
        if not acct.get("scan_folders"):
            acct["scan_folders"] = ["inbox"]

        # Channels default: all on if omitted. A missing key inside the block
        # also defaults to True.
        ch = acct.get("channels")
        if not isinstance(ch, dict):
            ch = {}
        ch = {
            "pushover": bool(ch.get("pushover", True)),
            "toast": bool(ch.get("toast", True)),
            "alexa": bool(ch.get("alexa", True)),
        }
        acct["channels"] = ch

        # Triggers list (may be empty -- account simply never fires).
        trigs = acct.get("triggers")
        if not isinstance(trigs, list):
            trigs = []
        acct["triggers"] = trigs

        # Per-provider required fields and placeholder detection.
        if provider == "microsoft":
            user_email = str(acct.get("user_email") or "").strip()
            if not user_email:
                raise ConfigError(
                    f"microsoft account '{name}' is missing 'user_email'."
                )
            if user_email.upper().startswith("YOUR_") or \
                    user_email in ("you@example.com", "you@outlook.com"):
                raise ConfigError(
                    f"microsoft account '{name}' still has a placeholder "
                    f"'user_email' ({user_email}). Replace it."
                )
        else:  # gmail
            user_email = str(acct.get("email") or "").strip()
            app_pw = str(acct.get("app_password") or "").strip()
            if not user_email:
                raise ConfigError(
                    f"gmail account '{name}' is missing 'email'."
                )
            if not app_pw:
                raise ConfigError(
                    f"gmail account '{name}' is missing 'app_password'."
                )
            if (user_email.upper().startswith("YOUR_")
                    or user_email in ("you@gmail.com",)
                    or app_pw.upper().startswith("YOUR_")
                    or app_pw in ("xxxx xxxx xxxx xxxx",)):
                raise ConfigError(
                    f"gmail account '{name}' still has placeholder "
                    "credentials. Replace 'email' and 'app_password'."
                )

    # microsoft client_id is required only if at least one microsoft account
    # exists.
    has_microsoft = any(a["provider"] == "microsoft" for a in accounts)
    if has_microsoft:
        cid = str(ms.get("client_id") or "").strip()
        if not cid or cid.upper().startswith("YOUR_"):
            raise ConfigError(
                "config.json has microsoft accounts but 'microsoft.client_id' "
                "is missing or still a placeholder. Register one public-client "
                "app in Microsoft Entra ID and paste its Application (client) "
                "ID here."
            )

    cfg["accounts"] = accounts

    # --- state directory -----------------------------------------------------
    default_appdata = os.environ.get("LOCALAPPDATA") or os.path.join(
        os.path.expanduser("~"), "AppData", "Local"
    )
    raw_state_dir = cfg.get("state_dir") or os.path.join(
        default_appdata, "UrgentEmailWatch"
    )
    state_dir = os.path.expandvars(os.path.expanduser(str(raw_state_dir)))
    try:
        os.makedirs(state_dir, exist_ok=True)
    except Exception as e:
        raise ConfigError(f"could not create state_dir '{state_dir}': {e}")
    cfg["_state_dir_resolved"] = state_dir

    return cfg


# --- Per-account state file paths --------------------------------------------

def _state_dir(cfg):
    return cfg["_state_dir_resolved"]


def _cursor_path(cfg, slug):
    return os.path.join(_state_dir(cfg), f"state_{slug}.json")


def _pending_path(cfg, slug):
    return os.path.join(_state_dir(cfg), f"pending_{slug}.json")


def _token_path(cfg, slug):
    return os.path.join(_state_dir(cfg), f"token_{slug}.json")


def _log_path(cfg):
    """The alert log is shared across accounts; entries carry the slug."""
    return os.path.join(_state_dir(cfg), "alert_log.json")


# --- Normalized email dict ---------------------------------------------------
#
# Every provider returns messages in this shape:
#   {
#     "id":            stable unique id for idempotency,
#     "subject":       str,
#     "from_address":  str,
#     "from_name":     str,
#     "body_preview":  str,
#     "received":      ISO-8601 str (best effort),
#     "is_read":       bool,
#     "folder":        the configured folder name this came from,
#   }


def _parse_from_header(raw_from):
    """Parse an RFC 2822 From header into (address, display_name)."""
    if not raw_from:
        return "", ""
    name, addr = email.utils.parseaddr(str(raw_from))
    return (addr or "").strip(), (name or "").strip()


# --- MailProvider abstraction -----------------------------------------------

class MailProvider:
    """Base interface every provider implements.

    A provider is created once per account and reused across poll passes. It
    owns whatever connection / token state that provider needs. All methods
    should swallow transient errors and either return an empty result or raise
    a clear exception on a hard failure -- the per-account pass is wrapped in a
    try/except so a raised exception is logged and never kills the loop.

    Subclasses must implement:
      connect()                        -- establish auth / connection
      list_new_messages(cursor)        -- return (messages, new_cursor)
      is_message_read(message_id)      -- bool
      send_mail(to_addr, subject, body)-- (ok: bool, detail: str)
    """

    #: Provider type string, e.g. "microsoft" / "gmail".
    kind = "base"

    def __init__(self, cfg, account):
        self.cfg = cfg
        self.account = account
        self.slug = account["_slug"]
        self.name = account["name"]

    def connect(self):
        """Establish auth / connectivity. Raise on a hard failure."""
        raise NotImplementedError

    def list_new_messages(self, cursor):
        """List messages newer than `cursor`.

        Returns a tuple (messages, new_cursor):
          messages    -- list of normalized email dicts (see above)
          new_cursor  -- the cursor value to persist for the next pass.
                         If nothing new arrived this should still be a valid
                         cursor (typically unchanged or advanced to "now").

        `cursor` is None on the very first run for this account.
        """
        raise NotImplementedError

    def is_message_read(self, message_id):
        """Return True if the given message is currently marked read."""
        raise NotImplementedError

    def send_mail(self, to_addr, subject, body):
        """Send an outbound plain-text email (used by the Pushover channel).
        Returns (ok: bool, detail: str)."""
        raise NotImplementedError

    # -- cursor helpers, shared by both providers -----------------------------

    def initial_cursor(self):
        """The cursor to use on the very first pass when no state exists.
        Time-based providers look back INITIAL_LOOKBACK_MIN minutes."""
        cutoff = datetime.now(timezone.utc) - timedelta(
            minutes=INITIAL_LOOKBACK_MIN
        )
        return cutoff.isoformat().replace("+00:00", "Z")


# --- Provider: Microsoft Graph (delegated device-code OAuth2) ----------------

class MicrosoftProvider(MailProvider):
    """Microsoft Graph provider using delegated device-code OAuth2.

    One public-client app registration (config: microsoft.client_id) serves
    every microsoft account. Each account holds a refresh token on disk
    (token_<slug>.json), acquired once via the device-code flow (--setup) and
    refreshed silently on every run.
    """

    kind = "microsoft"

    def __init__(self, cfg, account):
        super().__init__(cfg, account)
        self.client_id = cfg["microsoft"]["client_id"]
        self.authority = (
            cfg["microsoft"].get("authority") or DEFAULT_AUTHORITY
        ).rstrip("/")
        self.user_email = account["user_email"]
        # In-memory access-token cache for this account.
        self._access_token = None
        self._access_expires_at = 0.0
        # display-name -> folder id cache.
        self._folder_id_cache = {}

    # -- token file ----------------------------------------------------------

    def _read_token_file(self):
        return _load_json(_token_path(self.cfg, self.slug), None)

    def _write_refresh_token(self, refresh_token):
        _save_json(_token_path(self.cfg, self.slug), {
            "account": self.name,
            "user_email": self.user_email,
            "refresh_token": refresh_token,
            "saved_at": _now_iso(),
        })

    def has_valid_token(self):
        """True if a refresh token is on disk for this account."""
        data = self._read_token_file()
        return bool(data and data.get("refresh_token"))

    # -- device-code flow (interactive, used by --setup) ---------------------

    def device_code_setup(self):
        """Run the OAuth2 device-code flow interactively and save the
        resulting refresh token. Returns True on success.

        Step 1: POST /devicecode -> device_code, user_code, verification_uri,
                interval, expires_in, message.
        Step 2: print the message; poll /token until success or failure.
        """
        dc_url = f"{self.authority}/oauth2/v2.0/devicecode"
        body = urllib.parse.urlencode({
            "client_id": self.client_id,
            "scope": GRAPH_DELEGATED_SCOPE,
        }).encode("utf-8")
        req = urllib.request.Request(
            dc_url, data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                dc = json.loads(resp.read())
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="replace")
            _alog(self.slug,
                  f"device-code request failed: HTTP {e.code}: {detail[:300]}")
            return False
        except Exception as e:
            _alog(self.slug, f"device-code request error: {e}")
            return False

        device_code = dc.get("device_code")
        if not device_code:
            _alog(self.slug, f"device-code response missing device_code: {dc}")
            return False

        interval = int(dc.get("interval", 5)) or 5
        expires_in = int(dc.get("expires_in", DEVICE_CODE_MAX_WAIT_SEC))

        print()
        print("=" * 64)
        print(f"  Sign in for account: {self.name}  ({self.user_email})")
        print("-" * 64)
        # The authority hands back a human-readable instruction message.
        print("  " + (dc.get("message")
              or (f"To sign in, open {dc.get('verification_uri')} and "
                  f"enter the code {dc.get('user_code')}.")))
        print("=" * 64)
        print(f"  Waiting for you to complete sign-in (code expires in "
              f"~{expires_in // 60} min)...")
        print()

        token_url = f"{self.authority}/oauth2/v2.0/token"
        deadline = time.time() + min(expires_in, DEVICE_CODE_MAX_WAIT_SEC)

        while time.time() < deadline:
            time.sleep(interval)
            poll_body = urllib.parse.urlencode({
                "grant_type":
                    "urn:ietf:params:oauth:grant-type:device_code",
                "client_id": self.client_id,
                "device_code": device_code,
            }).encode("utf-8")
            poll_req = urllib.request.Request(
                token_url, data=poll_body,
                headers={"Content-Type": "application/x-www-form-urlencoded"},
                method="POST",
            )
            try:
                with urllib.request.urlopen(poll_req, timeout=20) as resp:
                    tok = json.loads(resp.read())
            except urllib.error.HTTPError as e:
                # The authority returns 400 with an OAuth error code while the
                # user has not finished signing in.
                try:
                    err = json.loads(e.read().decode("utf-8", "replace"))
                except Exception:
                    err = {}
                code = err.get("error", "")
                if code == "authorization_pending":
                    continue  # keep polling
                if code == "slow_down":
                    interval += 5  # back off as instructed
                    continue
                if code in ("expired_token", "code_expired"):
                    _alog(self.slug,
                          "sign-in code expired before completion. "
                          "Re-run 'python watchdog.py --setup'.")
                    return False
                if code == "authorization_declined":
                    _alog(self.slug,
                          "sign-in was declined. Re-run --setup to try again.")
                    return False
                _alog(self.slug,
                      f"device-code token error: {code or e.code}: "
                      f"{err.get('error_description', '')[:200]}")
                return False
            except Exception as e:
                _alog(self.slug, f"device-code poll error: {e}")
                continue  # transient -- keep trying until the deadline

            access_token = tok.get("access_token")
            refresh_token = tok.get("refresh_token")
            if access_token and refresh_token:
                self._access_token = access_token
                self._access_expires_at = (
                    time.time() + int(tok.get("expires_in", 3600))
                )
                self._write_refresh_token(refresh_token)
                _alog(self.slug, f"sign-in complete; refresh token saved to "
                      f"{_token_path(self.cfg, self.slug)}")
                return True
            # No error and no tokens -- unexpected; keep polling.

        _alog(self.slug,
              "timed out waiting for sign-in. Re-run --setup to try again.")
        return False

    # -- silent refresh ------------------------------------------------------

    def _refresh_access_token(self):
        """Exchange the saved refresh token for a fresh access token.

        Persists a rotated refresh token if the authority returns one.
        Raises RuntimeError on a hard failure so callers can surface it.
        """
        data = self._read_token_file()
        refresh_token = (data or {}).get("refresh_token")
        if not refresh_token:
            raise RuntimeError(
                f"no refresh token for account '{self.name}'. "
                "Run 'python watchdog.py --setup' once to sign in."
            )

        token_url = f"{self.authority}/oauth2/v2.0/token"
        body = urllib.parse.urlencode({
            "grant_type": "refresh_token",
            "client_id": self.client_id,
            "refresh_token": refresh_token,
            "scope": GRAPH_DELEGATED_SCOPE,
        }).encode("utf-8")
        req = urllib.request.Request(
            token_url, data=body,
            headers={"Content-Type": "application/x-www-form-urlencoded"},
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                tok = json.loads(resp.read())
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="replace")
            raise RuntimeError(
                f"token refresh failed for '{self.name}': HTTP {e.code}: "
                f"{detail[:300]} -- the refresh token may be revoked or "
                "expired; re-run 'python watchdog.py --setup'."
            )
        except Exception as e:
            raise RuntimeError(f"token refresh error for '{self.name}': {e}")

        access_token = tok.get("access_token")
        if not access_token:
            raise RuntimeError(
                f"token refresh response had no access_token: {tok}"
            )
        self._access_token = access_token
        self._access_expires_at = (
            time.time() + int(tok.get("expires_in", 3600))
        )
        # The authority rotates refresh tokens; persist the new one if given.
        rotated = tok.get("refresh_token")
        if rotated and rotated != refresh_token:
            self._write_refresh_token(rotated)
        return access_token

    def _token(self, force_refresh=False):
        """Return a valid access token, refreshing if near expiry."""
        now = time.time()
        if (not force_refresh
                and self._access_token
                and now < self._access_expires_at - TOKEN_REFRESH_MARGIN_SEC):
            return self._access_token
        return self._refresh_access_token()

    # -- MailProvider interface ----------------------------------------------

    def connect(self):
        """Establish a working delegated session by acquiring an access
        token from the saved refresh token."""
        self._token(force_refresh=True)
        _alog(self.slug, f"Microsoft Graph delegated token acquired "
              f"({self.user_email})")

    def _resolve_folder_id(self, display_name):
        """Look up a mail folder by displayName via the delegated /me/
        endpoint. Returns the id, or None if not found. (Cached.)"""
        cache_key = display_name.lower()
        if cache_key in self._folder_id_cache:
            return self._folder_id_cache[cache_key]

        token = self._token()
        # OData escaping: a single quote inside a literal is doubled.
        odata_value = display_name.replace("'", "''")
        params = {"$filter": f"displayName eq '{odata_value}'"}
        url = (f"{GRAPH_BASE}/me/mailFolders?"
               + urllib.parse.urlencode(params))
        req = urllib.request.Request(
            url, headers={"Authorization": f"Bearer {token}"})
        try:
            with urllib.request.urlopen(req, timeout=15) as resp:
                data = json.loads(resp.read())
        except Exception as e:
            _alog(self.slug, f"folder resolve error for '{display_name}': {e}")
            return None
        folders = data.get("value", [])
        if not folders:
            return None
        fid = folders[0].get("id")
        self._folder_id_cache[cache_key] = fid
        return fid

    def _list_folder(self, folder, since_iso):
        """List messages in one folder received after since_iso. Returns a
        list of normalized email dicts; [] on error (logged, not raised)."""
        token = self._token()
        if folder.lower() == "inbox":
            url_base = f"{GRAPH_BASE}/me/mailFolders/inbox/messages"
        else:
            fid = self._resolve_folder_id(folder)
            if not fid:
                _alog(self.slug, f"folder not found: {folder}")
                return []
            url_base = f"{GRAPH_BASE}/me/mailFolders/{fid}/messages"

        params = {
            "$filter": f"receivedDateTime gt {since_iso}",
            "$orderby": "receivedDateTime desc",
            "$top": "50",
            "$select": "id,subject,from,bodyPreview,receivedDateTime,isRead,internetMessageId",
        }
        url = url_base + "?" + urllib.parse.urlencode(params)
        headers = {
            "Authorization": f"Bearer {token}",
            "Prefer": 'outlook.body-content-type="text"',
        }
        # Graph returns at most $top results per page. After a long outage a
        # folder can hold >50 missed messages, so follow @odata.nextLink.
        # Cap the page count to bound the work.
        values = []
        next_url = url
        pages = 0
        max_pages = 20
        while next_url and pages < max_pages:
            pages += 1
            req = urllib.request.Request(next_url, headers=headers)
            try:
                with urllib.request.urlopen(req, timeout=20) as resp:
                    data = json.loads(resp.read())
            except urllib.error.HTTPError as e:
                detail = e.read().decode("utf-8", errors="replace")
                _alog(self.slug,
                      f"list messages failed in {folder}: HTTP {e.code} "
                      f"{detail[:200]}")
                return []
            except Exception as e:
                _alog(self.slug, f"list messages error in {folder}: {e}")
                return []
            values.extend(data.get("value", []))
            next_url = data.get("@odata.nextLink")

        out = []
        for m in values:
            sender = (m.get("from") or {}).get("emailAddress") or {}
            out.append({
                "id": m.get("id"),
                # Cross-folder dedup key: the RFC822 Message-ID is stable even
                # when a message moves between folders (which changes its
                # Graph id). Fall back to the Graph id if absent.
                "dedup_id": m.get("internetMessageId") or m.get("id"),
                "subject": m.get("subject") or "",
                "from_address": sender.get("address") or "",
                "from_name": sender.get("name") or "",
                "body_preview": m.get("bodyPreview") or "",
                "received": m.get("receivedDateTime") or "",
                "is_read": bool(m.get("isRead")),
                "folder": folder,
            })
        return out

    def list_new_messages(self, cursor):
        """List new messages across all configured folders. The cursor is the
        ISO timestamp of the last poll. The new cursor is 'now'."""
        # Capture the cursor timestamp BEFORE the folder queries. If we waited
        # until after, mail that arrived during the query window would fall
        # into the gap and be skipped by both this poll and the next.
        poll_started = _now_iso()
        since_iso = cursor or self.initial_cursor()
        messages = []
        for folder in self.account.get("scan_folders", ["inbox"]):
            messages.extend(self._list_folder(folder, since_iso))
        # Advance the time cursor to when this poll began.
        return messages, poll_started

    def is_message_read(self, message_id):
        """Delegated GET to check the isRead flag of one message."""
        try:
            token = self._token()
            url = (f"{GRAPH_BASE}/me/messages/{message_id}"
                   "?$select=isRead")
            req = urllib.request.Request(
                url, headers={"Authorization": f"Bearer {token}"})
            with urllib.request.urlopen(req, timeout=10) as resp:
                data = json.loads(resp.read())
                return bool(data.get("isRead"))
        except Exception:
            return False

    def send_mail(self, to_addr, subject, body):
        """Send a plain-text email via the delegated /me/sendMail endpoint."""
        if not to_addr:
            return False, "no recipient address"
        try:
            token = self._token()
        except Exception as e:
            return False, f"token error: {e}"

        payload = {
            "message": {
                "subject": subject,
                "body": {"contentType": "Text", "content": body},
                "toRecipients": [{"emailAddress": {"address": to_addr}}],
            },
            "saveToSentItems": False,  # don't pollute Sent Items
        }
        url = f"{GRAPH_BASE}/me/sendMail"
        req = urllib.request.Request(
            url, data=json.dumps(payload).encode("utf-8"),
            headers={
                "Authorization": f"Bearer {token}",
                "Content-Type": "application/json",
            },
            method="POST",
        )
        try:
            with urllib.request.urlopen(req, timeout=20) as resp:
                if 200 <= resp.status < 300:
                    return True, "push email accepted by Graph"
                return False, f"unexpected status {resp.status}"
        except urllib.error.HTTPError as e:
            detail = e.read().decode("utf-8", errors="replace")
            return False, f"HTTP {e.code}: {detail[:200]}"
        except Exception as e:
            return False, f"send error: {e}"


# --- Provider: Gmail (IMAP read + SMTP send, App Password) -------------------

class GmailProvider(MailProvider):
    """Gmail provider using IMAP for reading and SMTP for sending.

    Authentication is an App Password (generated after 2-Step Verification is
    on). New mail is tracked by IMAP UID: the cursor is the highest UID seen.
    """

    kind = "gmail"

    IMAP_HOST = "imap.gmail.com"
    IMAP_PORT = 993
    SMTP_HOST = "smtp.gmail.com"
    SMTP_PORT = 587  # STARTTLS

    def __init__(self, cfg, account):
        super().__init__(cfg, account)
        self.email_addr = account["email"]
        self.app_password = account["app_password"]

    # -- IMAP helpers --------------------------------------------------------

    def _imap_connect(self):
        """Open an authenticated IMAP4_SSL connection. Caller must logout."""
        imap = imaplib.IMAP4_SSL(self.IMAP_HOST, self.IMAP_PORT)
        imap.login(self.email_addr, self.app_password)
        return imap

    @staticmethod
    def _imap_logout(imap):
        try:
            imap.logout()
        except Exception:
            pass

    def initial_cursor(self):
        """Gmail's cursor is a per-folder map
        {folder: {"uidvalidity": <str>, "uid": <int>}}. On first run it is an
        empty dict; list_new_messages then primes each folder to the current
        high watermark so historical mail is not re-alerted."""
        return {}

    def connect(self):
        """Verify the IMAP login works and the SMTP login works."""
        imap = self._imap_connect()
        self._imap_logout(imap)
        # Verify SMTP auth too, so --test catches a bad App Password early.
        smtp = self._smtp_connect()
        try:
            smtp.quit()
        except Exception:
            pass
        _alog(self.slug, f"Gmail IMAP+SMTP login verified ({self.email_addr})")

    def _mailbox_name(self, folder):
        """Map a configured folder name to the IMAP mailbox name. 'inbox'
        maps to INBOX; other names are passed through (Gmail labels are
        addressable as folder names). Labels with spaces are quoted."""
        mailbox = "INBOX" if folder.lower() == "inbox" else folder
        if " " in mailbox and not mailbox.startswith('"'):
            mailbox = '"' + mailbox + '"'
        return mailbox

    def _select_folder(self, imap, folder):
        """SELECT a folder/label. 'inbox' maps to INBOX; other names are
        passed through (Gmail labels are addressable as folder names).
        Returns True on success."""
        mailbox = self._mailbox_name(folder)
        try:
            typ, _ = imap.select(mailbox, readonly=True)
            return typ == "OK"
        except Exception as e:
            _alog(self.slug, f"IMAP select failed for '{folder}': {e}")
            return False

    def _uidvalidity(self, imap, folder):
        """Return the current UIDVALIDITY for `folder` as a string, or None
        if it cannot be determined. Reads the value the SELECT response left
        in untagged_responses, falling back to an explicit STATUS query."""
        try:
            resp = imap.untagged_responses.get("UIDVALIDITY")
            if resp:
                val = resp[0]
                if isinstance(val, bytes):
                    val = val.decode("ascii", "replace")
                return str(val).strip()
        except Exception:
            pass
        try:
            mailbox = self._mailbox_name(folder)
            typ, data = imap.status(mailbox, "(UIDVALIDITY)")
            if typ == "OK" and data and data[0]:
                blob = data[0]
                if isinstance(blob, bytes):
                    blob = blob.decode("ascii", "replace")
                m = re.search(r"UIDVALIDITY\s+(\d+)", blob)
                if m:
                    return m.group(1)
        except Exception as e:
            _alog(self.slug,
                  f"IMAP UIDVALIDITY lookup failed for '{folder}': {e}")
        return None

    @staticmethod
    def _decode_header(raw):
        """Decode an RFC 2047 encoded header into a plain string."""
        if raw is None:
            return ""
        try:
            parts = email.header.decode_header(raw)
            out = []
            for text, enc in parts:
                if isinstance(text, bytes):
                    out.append(text.decode(enc or "utf-8", errors="replace"))
                else:
                    out.append(text)
            return "".join(out)
        except Exception:
            return str(raw)

    @staticmethod
    def _body_snippet(msg, limit=600):
        """Extract a short plain-text snippet from an email.message.Message."""
        try:
            if msg.is_multipart():
                for part in msg.walk():
                    if part.get_content_type() == "text/plain" and \
                            "attachment" not in str(
                                part.get("Content-Disposition", "")):
                        payload = part.get_payload(decode=True)
                        if payload:
                            charset = part.get_content_charset() or "utf-8"
                            text = payload.decode(charset, errors="replace")
                            return " ".join(text.split())[:limit]
                return ""
            payload = msg.get_payload(decode=True)
            if payload:
                charset = msg.get_content_charset() or "utf-8"
                text = payload.decode(charset, errors="replace")
                return " ".join(text.split())[:limit]
        except Exception:
            pass
        return ""

    def _fetch_folder(self, imap, folder, cursor_entry):
        """Fetch messages in `folder` with UID greater than the cursor.

        `cursor_entry` is this folder's stored cursor value, expected to be a
        dict {"uidvalidity": <str>, "uid": <int>}. It may also be None (never
        seen), or an int from the legacy schema.

        Returns (messages, new_entry) where new_entry is the cursor value to
        persist for this folder. new_entry is None when the folder could not
        be primed -- the absent entry makes the folder re-prime next poll
        rather than treating the whole mailbox as new (Fix A).

        First sight of a folder -- no entry, a legacy int entry, or a changed
        UIDVALIDITY -- primes to the current max UID and alerts nothing
        (Minor 2 / UIDVALIDITY safety + legacy-cursor migration).
        """
        if not self._select_folder(imap, folder):
            # Could not select -- preserve the existing cursor entry as-is
            # (when it is a dict) so the folder is reconsidered next poll.
            return [], cursor_entry if isinstance(cursor_entry, dict) else None

        cur_uidvalidity = self._uidvalidity(imap, folder)

        # Determine last_uid only if the stored entry is a dict whose
        # uidvalidity matches the folder's current UIDVALIDITY. Anything else
        # (no entry, legacy int entry, or a changed UIDVALIDITY) is treated as
        # first sight and the folder is primed.
        last_uid = None
        if (isinstance(cursor_entry, dict)
                and cur_uidvalidity is not None
                and str(cursor_entry.get("uidvalidity")) == cur_uidvalidity):
            try:
                last_uid = int(cursor_entry.get("uid"))
            except (TypeError, ValueError):
                last_uid = None

        if last_uid is None:
            # Prime: record the current high-water UID, alert nothing.
            try:
                typ, data = imap.uid("search", None, "ALL")
                if typ != "OK":
                    # Prime failed -- leave the folder unprimed so it
                    # re-primes next poll instead of persisting a 0 cursor
                    # that would treat the whole mailbox as new (Fix A).
                    return [], None
                uids = data[0].split() if data and data[0] else []
                high = max((int(u) for u in uids), default=0)
                if cur_uidvalidity is None:
                    # Without a UIDVALIDITY we cannot store a safe dict
                    # entry; leave unprimed so we retry next poll (Fix A).
                    return [], None
                return [], {"uidvalidity": cur_uidvalidity, "uid": high}
            except Exception as e:
                _alog(self.slug, f"IMAP prime failed for '{folder}': {e}")
                # Prime failed -- leave the folder unprimed (Fix A).
                return [], None

        # From here on we have a valid last_uid and a matching UIDVALIDITY.
        # The new cursor entry, if unchanged, keeps the same UIDVALIDITY.
        unchanged_entry = {"uidvalidity": cur_uidvalidity, "uid": last_uid}

        # Search for UIDs strictly greater than the cursor.
        try:
            typ, data = imap.uid("search", None, f"UID {last_uid + 1}:*")
            if typ != "OK":
                return [], unchanged_entry
            uids = data[0].split() if data and data[0] else []
        except Exception as e:
            _alog(self.slug, f"IMAP search failed for '{folder}': {e}")
            return [], unchanged_entry

        # "UID n:*" can echo back the cursor UID itself; filter to > last_uid.
        new_uids = sorted(
            u for u in (int(x) for x in uids) if u > last_uid
        )
        if not new_uids:
            return [], unchanged_entry

        messages = []
        # high advances only across the contiguous prefix of UIDs that were
        # successfully fetched AND parsed. The first failure freezes the
        # high-water mark so the failed UID -- and everything after it -- is
        # re-examined next poll; later messages are still processed and
        # appended, and idempotency dedups any re-alerts (Fix C).
        high = last_uid
        contiguous_ok = True
        for uid in new_uids:
            ok = False
            try:
                typ, fetched = imap.uid(
                    "fetch", str(uid),
                    "(FLAGS BODY.PEEK[HEADER] BODY.PEEK[TEXT])")
                if typ != "OK" or not fetched:
                    contiguous_ok = False
                    continue
            except Exception as e:
                _alog(self.slug, f"IMAP fetch failed for UID {uid}: {e}")
                contiguous_ok = False
                continue

            flags = b""
            header_bytes = b""
            text_bytes = b""
            for item in fetched:
                if isinstance(item, tuple) and len(item) == 2:
                    meta, payload = item
                    meta_s = meta.decode("utf-8", "replace") if isinstance(
                        meta, bytes) else str(meta)
                    if "HEADER" in meta_s:
                        header_bytes = payload or b""
                    elif "TEXT" in meta_s:
                        text_bytes = payload or b""
                    if "FLAGS" in meta_s:
                        flags += meta if isinstance(meta, bytes) else b""
                elif isinstance(item, bytes):
                    flags += item

            try:
                hdr_msg = email.message_from_bytes(header_bytes)
            except Exception:
                hdr_msg = email.message.Message()

            subject = self._decode_header(hdr_msg.get("Subject"))
            from_raw = self._decode_header(hdr_msg.get("From"))
            from_addr, from_name = _parse_from_header(from_raw)
            message_id = (hdr_msg.get("Message-ID") or "").strip()
            date_hdr = hdr_msg.get("Date")
            received_iso = ""
            if date_hdr:
                try:
                    dt = email.utils.parsedate_to_datetime(date_hdr)
                    if dt is not None:
                        if dt.tzinfo is None:
                            dt = dt.replace(tzinfo=timezone.utc)
                        received_iso = dt.astimezone(
                            timezone.utc).isoformat().replace(
                            "+00:00", "Z")
                except Exception:
                    received_iso = ""

            # Body snippet: parse the TEXT part as a standalone message.
            body_preview = ""
            if text_bytes:
                try:
                    body_msg = email.message_from_bytes(
                        header_bytes + b"\r\n" + text_bytes)
                    body_preview = self._body_snippet(body_msg)
                except Exception:
                    body_preview = " ".join(
                        text_bytes.decode("utf-8", "replace").split()
                    )[:600]

            is_read = b"\\Seen" in flags

            # Stable id: prefer the RFC822 Message-ID; fall back to a
            # folder+UID composite so it stays unique within the account.
            stable_id = message_id or f"{folder}:UID{uid}"

            messages.append({
                "id": stable_id,
                # Cross-folder dedup key: prefer the RFC822 Message-ID, which
                # is stable when a message moves between folders. Fall back to
                # the folder:UID composite (the existing stable id).
                "dedup_id": message_id or f"{folder}:UID{uid}",
                "_uid": uid,           # internal: used by is_message_read
                "_folder_raw": folder,
                "subject": subject,
                "from_address": from_addr,
                "from_name": from_name,
                "body_preview": body_preview,
                "received": received_iso,
                "is_read": is_read,
                "folder": folder,
            })

            # This UID was fetched AND parsed AND appended. Advance the
            # high-water mark only while the success run is still unbroken
            # (Fix C: the first failure above freezes high here).
            ok = True
            if contiguous_ok and ok:
                high = uid

        return messages, {"uidvalidity": cur_uidvalidity, "uid": high}

    def list_new_messages(self, cursor):
        """Walk every configured folder, returning new messages and the
        updated per-folder cursor map.

        The cursor map keys folders to {"uidvalidity": <str>, "uid": <int>}.
        A folder whose new entry is None (prime failed) is left OUT of the
        returned map so it re-primes on the next poll (Fix A).
        """
        old_map = dict(cursor) if isinstance(cursor, dict) else {}
        new_map = {}
        messages = []
        imap = None
        try:
            imap = self._imap_connect()
            for folder in self.account.get("scan_folders", ["inbox"]):
                # cursor_entry may be a dict (current schema), an int (legacy
                # schema), or absent (first sight).
                cursor_entry = old_map.get(folder)
                folder_msgs, new_entry = self._fetch_folder(
                    imap, folder, cursor_entry)
                messages.extend(folder_msgs)
                if new_entry is not None:
                    new_map[folder] = new_entry
                # new_entry is None => omit the folder so it re-primes.
        finally:
            if imap is not None:
                self._imap_logout(imap)
        return messages, new_map

    def is_message_read(self, message_id):
        """Re-fetch FLAGS for the message and look for the \\Seen flag.

        message_id is the stable id we assigned. It encodes the folder and UID
        when it has the 'folder:UIDn' shape; for RFC822-Message-ID ids we fall
        back to an IMAP HEADER search across folders.
        """
        imap = None
        try:
            imap = self._imap_connect()
            # Composite id form: "<folder>:UID<n>".
            m = re.match(r"^(.*):UID(\d+)$", message_id or "")
            if m:
                folder, uid = m.group(1), m.group(2)
                if not self._select_folder(imap, folder):
                    return False
                typ, data = imap.uid("fetch", uid, "(FLAGS)")
                if typ != "OK" or not data:
                    return False
                blob = b" ".join(
                    x if isinstance(x, bytes) else b"" for x in data)
                return b"\\Seen" in blob
            # RFC822 Message-ID form: search each folder by header.
            for folder in self.account.get("scan_folders", ["inbox"]):
                if not self._select_folder(imap, folder):
                    continue
                typ, data = imap.uid(
                    "search", None, "HEADER", "Message-ID",
                    '"' + message_id + '"')
                if typ != "OK" or not data or not data[0]:
                    continue
                uid = data[0].split()[0].decode()
                typ, fdata = imap.uid("fetch", uid, "(FLAGS)")
                if typ != "OK" or not fdata:
                    continue
                blob = b" ".join(
                    x if isinstance(x, bytes) else b"" for x in fdata)
                return b"\\Seen" in blob
            return False
        except Exception:
            return False
        finally:
            if imap is not None:
                self._imap_logout(imap)

    def _smtp_connect(self):
        """Open an authenticated SMTP connection (STARTTLS on 587)."""
        ctx = ssl.create_default_context()
        smtp = smtplib.SMTP(self.SMTP_HOST, self.SMTP_PORT, timeout=20)
        smtp.ehlo()
        smtp.starttls(context=ctx)
        smtp.ehlo()
        smtp.login(self.email_addr, self.app_password)
        return smtp

    def send_mail(self, to_addr, subject, body):
        """Send a plain-text email via Gmail SMTP."""
        if not to_addr:
            return False, "no recipient address"
        msg = email.message.EmailMessage()
        msg["Subject"] = subject
        msg["From"] = self.email_addr
        msg["To"] = to_addr
        msg.set_content(body)
        smtp = None
        try:
            smtp = self._smtp_connect()
            smtp.send_message(msg)
            return True, "push email sent via Gmail SMTP"
        except Exception as e:
            return False, f"SMTP send error: {e}"
        finally:
            if smtp is not None:
                try:
                    smtp.quit()
                except Exception:
                    pass


# --- Provider factory --------------------------------------------------------

def make_provider(cfg, account):
    """Construct the MailProvider for an account based on its provider type."""
    provider = account["provider"]
    if provider == "microsoft":
        return MicrosoftProvider(cfg, account)
    if provider == "gmail":
        return GmailProvider(cfg, account)
    raise ConfigError(f"unknown provider '{provider}'")


# --- Trigger matching --------------------------------------------------------

def enabled_triggers(account):
    """Return the account's enabled trigger dicts."""
    return [t for t in account.get("triggers", []) if t.get("enabled", True)]


def match_email_to_triggers(email_dict, triggers):
    """Return the FIRST matching trigger dict, or None.

    A trigger matches if ANY of these conditions is true:
      - sender_domains: from_address ends with @<domain> or .<domain>
      - sender_addresses: from_address equals an entry (case-insensitive)
      - subject_keywords: a keyword is a substring of the subject
      - body_keywords: a keyword is a substring of the body preview
    All comparisons are case-insensitive.
    """
    from_addr = (email_dict.get("from_address") or "").lower()
    subject = (email_dict.get("subject") or "").lower()
    body = (email_dict.get("body_preview") or "").lower()

    for trig in triggers:
        for domain in trig.get("sender_domains", []):
            d = domain.lower().lstrip("@")
            if from_addr.endswith("@" + d) or from_addr.endswith("." + d):
                return trig
        for addr in trig.get("sender_addresses", []):
            if from_addr == addr.lower():
                return trig
        for kw in trig.get("subject_keywords", []):
            if kw and kw.lower() in subject:
                return trig
        for kw in trig.get("body_keywords", []):
            if kw and kw.lower() in body:
                return trig
    return None


def resolve_channels(account, trigger):
    """Decide which channels a matched email should fire.

    Precedence: a trigger MAY carry its own `channels` block to override the
    account default; if absent the trigger inherits the account's `channels`.
    A channel inside a block defaults to True when its key is missing.
    """
    acct_channels = account.get("channels", {})
    trig_channels = trigger.get("channels")
    if isinstance(trig_channels, dict):
        return {
            "pushover": bool(trig_channels.get("pushover", True)),
            "toast": bool(trig_channels.get("toast", True)),
            "alexa": bool(trig_channels.get("alexa", True)),
        }
    return {
        "pushover": bool(acct_channels.get("pushover", True)),
        "toast": bool(acct_channels.get("toast", True)),
        "alexa": bool(acct_channels.get("alexa", True)),
    }


# --- Channel: Alexa Notify Me ------------------------------------------------

def send_alexa_notify(cfg, notification_text):
    """POST a notification to the user's Echo devices via the Notify Me skill.

    Echo devices flash a yellow ring and queue the notification. The user says
    "Alexa, what are my notifications?" to hear them read aloud.
    Returns (success: bool, detail: str).
    """
    access_code = (
        (cfg.get("notifications") or {}).get("notify_me") or {}
    ).get("access_code", "")
    if not access_code:
        return False, "notify_me access_code not configured"

    payload = json.dumps({
        "accessCode": access_code,
        "notification": notification_text,
    }).encode("utf-8")
    req = urllib.request.Request(
        NOTIFY_ME_URL, data=payload,
        headers={"Content-Type": "application/json"}, method="POST",
    )
    try:
        with urllib.request.urlopen(req, timeout=10) as resp:
            if 200 <= resp.status < 300:
                return True, f"queued ({len(notification_text)} chars)"
            return False, f"unexpected status {resp.status}"
    except urllib.error.HTTPError as e:
        detail = e.read().decode("utf-8", errors="replace")
        return False, f"HTTP {e.code}: {detail[:200]}"
    except Exception as e:
        return False, f"send error: {e}"


def _build_alexa_text(account, email_dict, trigger):
    """Build the Notify Me message text, optimized for spoken delivery and
    capped at NOTIFY_ME_BODY_MAX characters."""
    trig_name = trigger.get("name", "Alert")
    from_name = (email_dict.get("from_name")
                 or email_dict.get("from_address") or "unknown sender")
    subject = email_dict.get("subject") or "no subject"
    text = (f"Urgent: {trig_name}. Account {account['name']}. "
            f"Email from {from_name}. Subject: {subject}")
    if len(text) > NOTIFY_ME_BODY_MAX:
        text = text[: NOTIFY_ME_BODY_MAX - 3] + "..."
    return text


# --- Channel: Windows toast --------------------------------------------------

def fire_toast(title, message):
    """Render a Windows toast by invoking the bundled Show-Toast.ps1.

    Returns (success: bool, detail: str). The PowerShell process is launched
    detached; we do not block on it.
    """
    if not os.path.exists(TOAST_SCRIPT):
        return False, f"Show-Toast.ps1 not found at {TOAST_SCRIPT}"
    try:
        subprocess.Popen(
            [
                "powershell", "-NoProfile", "-ExecutionPolicy", "Bypass",
                "-File", TOAST_SCRIPT,
                "-Title", title,
                "-Message", message,
            ],
            creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
            close_fds=True,
        )
    except FileNotFoundError:
        return False, "powershell executable not found on PATH"
    except Exception as e:
        return False, f"toast launch failed: {e}"
    return True, "toast launched"


# --- Alert message builders --------------------------------------------------

def _build_push_title(account, trigger):
    """Push notification title: 'URGENT [account]: <trigger name>'."""
    return f"URGENT [{account['name']}]: {trigger.get('name', 'Alert')}"


def _build_push_body(email_dict, trigger):
    """Build the push notification body. Format:
      From: FROM_NAME <ADDRESS>
      Subject: SUBJECT

      BODY_PREVIEW (truncated to fit)
    """
    from_name = email_dict.get("from_name") or ""
    from_addr = email_dict.get("from_address") or "?"
    from_line = (f"From: {from_name} <{from_addr}>" if from_name
                 else f"From: {from_addr}")
    subject = email_dict.get("subject") or "(no subject)"
    body = email_dict.get("body_preview") or ""

    composed = f"{from_line}\nSubject: {subject}"
    remaining = PUSH_BODY_MAX - len(composed) - 5
    if remaining > 50 and body:
        composed += "\n\n" + body[:remaining]
    return composed


def _build_toast_payload(account, email_dict, trigger):
    """Build the toast title and message strings."""
    title = f"URGENT [{account['name']}]: {trigger.get('name', 'Alert')}"
    from_name = (email_dict.get("from_name")
                 or email_dict.get("from_address") or "?")
    subject = email_dict.get("subject") or "(no subject)"
    message = f"From: {from_name}\n{subject}"
    return title, message


# --- Alert dispatch ----------------------------------------------------------

def dispatch_alert(cfg, account, provider, email_dict, trigger, fire_count=1):
    """Fire all eligible channels for this email + trigger.

    A channel fires only if it is enabled in the resolved channel set (account
    default, optionally overridden by the trigger) AND the global destination
    in `notifications` is configured. Returns a result dict for the audit log.
    """
    notif = cfg.get("notifications", {})
    channels = resolve_channels(account, trigger)
    result = {
        "account": account["name"],
        "account_slug": account["_slug"],
        "alert_id": email_dict["id"],
        "trigger_name": trigger.get("name"),
        "fire_count": fire_count,
        "fired_at": _now_iso(),
        "email": {
            "subject": email_dict.get("subject"),
            "from": email_dict.get("from_address"),
            "received": email_dict.get("received"),
            "folder": email_dict.get("folder"),
        },
        "channels": {},
    }
    subject_preview = (email_dict.get("subject", "") or "")[:50]
    slug = account["_slug"]

    # --- Pushover (via the provider's outbound send_mail) -------------------
    gateway = (notif.get("pushover") or {}).get("gateway_email", "")
    if channels["pushover"]:
        if not gateway:
            result["channels"]["pushover"] = {
                "ok": False,
                "detail": "pushover gateway_email not configured",
            }
        else:
            body = _build_push_body(email_dict, trigger)
            if len(body) > PUSH_BODY_MAX:
                body = body[: PUSH_BODY_MAX - 3] + "..."
            title = _build_push_title(account, trigger)
            ok, detail = provider.send_mail(gateway, title, body)
            result["channels"]["pushover"] = {"ok": ok, "detail": detail}
            if ok:
                _alog(slug, f"Pushover sent for '{subject_preview}' "
                      f"({detail})")
            else:
                _alog(slug, f"Pushover FAILED for '{subject_preview}': "
                      f"{detail}")

    # --- Toast --------------------------------------------------------------
    toast_enabled = bool((notif.get("toast") or {}).get("enabled", True))
    if channels["toast"]:
        if not toast_enabled:
            result["channels"]["toast"] = {
                "ok": False, "detail": "toast disabled in notifications",
            }
        else:
            title, message = _build_toast_payload(account, email_dict, trigger)
            ok, detail = fire_toast(title, message)
            result["channels"]["toast"] = {"ok": ok, "detail": detail}
            if ok:
                _alog(slug, f"Toast launched for '{subject_preview}'")
            else:
                _alog(slug, f"Toast FAILED for '{subject_preview}': {detail}")

    # --- Alexa --------------------------------------------------------------
    access_code = (notif.get("notify_me") or {}).get("access_code", "")
    if channels["alexa"]:
        if not access_code:
            result["channels"]["alexa"] = {
                "ok": False,
                "detail": "notify_me access_code not configured",
            }
        else:
            text = _build_alexa_text(account, email_dict, trigger)
            ok, detail = send_alexa_notify(cfg, text)
            result["channels"]["alexa"] = {"ok": ok, "detail": detail}
            if ok:
                _alog(slug, f"Alexa notify queued for '{subject_preview}' "
                      f"({detail})")
            else:
                _alog(slug, f"Alexa notify FAILED for '{subject_preview}': "
                      f"{detail}")

    return result


# --- Alert log (shared) ------------------------------------------------------

def _log_alert(cfg, record):
    """Append an alert record to the shared alert_log.json (last 500)."""
    log_file = _log_path(cfg)
    with _state_lock:
        log = _load_json(log_file, {"entries": []})
        log["entries"].append(record)
        log["entries"] = log["entries"][-500:]
        _save_json(log_file, log)


# --- Per-account cursor + pending state --------------------------------------

def _load_cursor(cfg, slug):
    """Load the poll cursor for one account. Returns None if there is none."""
    data = _load_json(_cursor_path(cfg, slug), {})
    return data.get("cursor") if isinstance(data, dict) else None


def _save_cursor(cfg, slug, cursor):
    """Persist the poll cursor for one account."""
    _save_json(_cursor_path(cfg, slug),
               {"cursor": cursor, "updated_at": _now_iso()})


def _load_pending(cfg, slug):
    """Load one account's pending-alerts file with its default shape."""
    return _load_json(_pending_path(cfg, slug), {"alerts": {}})


def _save_pending(cfg, slug, pending):
    """Persist one account's pending-alerts file."""
    _save_json(_pending_path(cfg, slug), pending)


# --- Pending alert checker (15-min re-fire), per account ---------------------

def check_pending_alerts(cfg, account, provider):
    """For one account: for each pending alert, if the email has been read,
    acknowledge and drop it; if it has been unread for more than 15 minutes
    and has not yet been re-fired, re-fire the Pushover channel once."""
    slug = account["_slug"]
    with _state_lock:
        pending = _load_pending(cfg, slug)

    now = datetime.now(timezone.utc)
    changed = False
    gateway = (
        (cfg.get("notifications") or {}).get("pushover") or {}
    ).get("gateway_email", "")

    for alert_id in list(pending["alerts"].keys()):
        entry = pending["alerts"][alert_id]

        if entry.get("acknowledged"):
            del pending["alerts"][alert_id]
            changed = True
            continue

        if entry.get("re_fire_count", 0) >= RE_FIRE_MAX_COUNT:
            continue  # already re-fired; stop tracking

        first_fired_iso = entry.get("first_fired_at")
        if not first_fired_iso:
            continue
        try:
            first_fired = datetime.fromisoformat(
                first_fired_iso.replace("Z", "+00:00"))
        except Exception:
            continue
        if (now - first_fired).total_seconds() < RE_FIRE_AFTER_SEC:
            continue

        # Has the email been read since the first fire?
        try:
            is_read = provider.is_message_read(alert_id)
        except Exception as e:
            _alog(slug, f"pending read-check error for "
                  f"{str(alert_id)[:24]}: {e}")
            continue
        if is_read:
            entry["acknowledged"] = True
            entry["acknowledged_at"] = _now_iso()
            del pending["alerts"][alert_id]
            changed = True
            _alog(slug, f"alert {str(alert_id)[:24]} acknowledged "
                  "(email read)")
            continue

        # Re-fire the Pushover channel only (toast may still be on screen).
        email_meta = entry.get("email", {})
        trigger = entry.get("trigger", {})
        re_email = {
            "from_name": email_meta.get("from"),
            "from_address": email_meta.get("from"),
            "subject": "[RE-ALERT] " + (email_meta.get("subject") or ""),
            "body_preview": "Email not yet read after 15 minutes.",
        }
        re_trigger = {"name": trigger.get("name", "ALERT")}
        if gateway:
            body = _build_push_body(re_email, re_trigger)
            if len(body) > PUSH_BODY_MAX:
                body = body[: PUSH_BODY_MAX - 3] + "..."
            title = f"RE-ALERT [{account['name']}]: {re_trigger['name']}"
            try:
                ok, detail = provider.send_mail(gateway, title, body)
            except Exception as e:
                ok, detail = False, f"send error: {e}"
        else:
            ok, detail = False, "pushover gateway_email not configured"
        entry["re_fire_count"] = entry.get("re_fire_count", 0) + 1
        entry["last_re_fire_at"] = _now_iso()
        entry["last_re_fire_result"] = {"ok": ok, "detail": detail}
        changed = True
        _alog(slug, f"RE-FIRE for "
              f"'{(email_meta.get('subject') or '')[:50]}': ok={ok}")
        # Once the final re-fire is done the entry would otherwise be skipped
        # forever yet kept on disk, so pending_<slug>.json grows without
        # bound. Drop it now -- the shared alert log still dedups the
        # message, so removing it from pending is safe (Fix D).
        if entry["re_fire_count"] >= RE_FIRE_MAX_COUNT:
            del pending["alerts"][alert_id]

    if changed:
        with _state_lock:
            _save_pending(cfg, slug, pending)


# --- One account's poll pass -------------------------------------------------

def run_account_pass(cfg, account, provider):
    """Run one scan pass for a single account: list new mail, match against
    that account's triggers, dispatch alerts. Returns the count of alerts
    fired. Exceptions propagate to the caller, which isolates them."""
    slug = account["_slug"]
    triggers = enabled_triggers(account)
    if not triggers:
        _alog(slug, "no enabled triggers; skipping pass")
        return 0

    cursor = _load_cursor(cfg, slug)
    if cursor is None:
        cursor = provider.initial_cursor()

    messages, new_cursor = provider.list_new_messages(cursor)

    pending = _load_pending(cfg, slug)
    already_pending = set(pending.get("alerts", {}).keys())

    log = _load_json(_log_path(cfg), {"entries": []})
    historical_ids = {
        e["alert_id"] for e in log.get("entries", [])
        if e.get("alert_id") and e.get("account_slug") == slug
    }

    seen_in_pass = set()
    fired = 0

    for email_dict in messages:
        mid = email_dict.get("id")
        # Within-pass dedup keys on dedup_id (stable across folders) so a
        # message that moved between two scanned folders is not double-
        # alerted. Pending/log/idempotency keying still uses `id` unchanged.
        dedup_key = email_dict.get("dedup_id") or email_dict.get("id")
        if not mid or dedup_key in seen_in_pass:
            continue
        seen_in_pass.add(dedup_key)

        # Idempotency: per-account namespace. A message already handled for
        # this account never fires again.
        if mid in already_pending or mid in historical_ids:
            continue

        trig = match_email_to_triggers(email_dict, triggers)
        if not trig:
            continue

        result = dispatch_alert(cfg, account, provider, email_dict, trig,
                                fire_count=1)
        _log_alert(cfg, result)

        with _state_lock:
            pending["alerts"][mid] = {
                "first_fired_at": result["fired_at"],
                "trigger": {"name": trig.get("name")},
                "email": result["email"],
                "re_fire_count": 0,
            }
            _save_pending(cfg, slug, pending)
        fired += 1

    # Advance the cursor only after a successful pass.
    _save_cursor(cfg, slug, new_cursor)
    return fired


# --- Provider registry (built once, reused across passes) --------------------

def build_providers(cfg):
    """Construct a MailProvider for every account. Returns a list of
    (account, provider) tuples."""
    pairs = []
    for account in cfg["accounts"]:
        pairs.append((account, make_provider(cfg, account)))
    return pairs


# --- Multi-account poll pass -------------------------------------------------

def run_all_accounts(cfg, providers, do_pending=False):
    """Run one poll pass across every account. An exception in one account is
    logged and never blocks the others. Returns total alerts fired."""
    total = 0
    for account, provider in providers:
        slug = account["_slug"]
        try:
            fired = run_account_pass(cfg, account, provider)
            total += fired
            if fired:
                _alog(slug, f"pass complete, {fired} alert(s) fired")
        except Exception as e:
            _alog(slug, f"account pass failed: {e}")
            import traceback
            _alog(slug, traceback.format_exc())
        if do_pending:
            try:
                check_pending_alerts(cfg, account, provider)
            except Exception as e:
                _alog(slug, f"pending check failed: {e}")
    return total


# --- Forever loop ------------------------------------------------------------

def watchdog_loop(cfg, providers):
    """Run forever, polling every poll_interval_sec. Resilient: an exception
    inside a pass is logged but never kills the loop."""
    interval = int(cfg.get("poll_interval_sec", 60))
    _log(f"watchdog started, {len(providers)} account(s), polling every "
         f"{interval}s")
    last_pending_check = 0.0

    while True:
        try:
            now = time.time()
            do_pending = (
                now - last_pending_check >= PENDING_CHECK_INTERVAL_SEC
            )
            total = run_all_accounts(cfg, providers, do_pending=do_pending)
            if do_pending:
                last_pending_check = now
            if total:
                _log(f"cycle complete, {total} alert(s) fired across "
                     "all accounts")
        except Exception as e:
            _log(f"unhandled error in cycle: {e}")
            import traceback
            _log(traceback.format_exc())
        time.sleep(interval)


# --- Command-line modes ------------------------------------------------------

def _mode_setup(cfg):
    """--setup: for every microsoft account without a valid cached refresh
    token, run the device-code sign-in. Gmail accounts need no setup."""
    print("urgent-email-watch v2 -- setup mode")
    print("=" * 60)

    ms_accounts = [a for a in cfg["accounts"] if a["provider"] == "microsoft"]
    gmail_accounts = [a for a in cfg["accounts"] if a["provider"] == "gmail"]

    if gmail_accounts:
        print(f"{len(gmail_accounts)} gmail account(s) need no interactive "
              "setup (App Password auth):")
        for a in gmail_accounts:
            print(f"  - {a['name']}")
        print()

    if not ms_accounts:
        print("No microsoft accounts -- nothing to set up.")
        return 0

    failures = 0
    for account in ms_accounts:
        provider = MicrosoftProvider(cfg, account)
        if provider.has_valid_token():
            # Confirm the saved token still works with a silent refresh.
            try:
                provider.connect()
                print(f"[{account['name']}] already signed in "
                      "(token valid). Skipping.")
                continue
            except Exception as e:
                print(f"[{account['name']}] saved token no longer works "
                      f"({e}). Re-running sign-in...")
        print(f"[{account['name']}] needs sign-in.")
        ok = provider.device_code_setup()
        if not ok:
            failures += 1

    print()
    if failures:
        print(f"Setup finished with {failures} account(s) NOT signed in. "
              "Re-run 'python watchdog.py --setup' to retry.")
        return 1
    print("Setup complete. All microsoft accounts are signed in.")
    print("Run 'python watchdog.py --test' to verify everything.")
    return 0


def _mode_test(cfg):
    """--test: validate config; for each account verify auth/connection works
    and list that account's triggers; exit."""
    print("urgent-email-watch v2 -- test mode")
    print("=" * 60)
    print(f"Config file: {CONFIG_FILE}")
    print(f"State dir:   {cfg['_state_dir_resolved']}")
    print(f"Poll interval: {cfg.get('poll_interval_sec')}s")
    print(f"Accounts:    {len(cfg['accounts'])}")
    print()

    notif = cfg.get("notifications", {})
    dests = []
    if (notif.get("pushover") or {}).get("gateway_email"):
        dests.append("pushover")
    if (notif.get("toast") or {}).get("enabled", True):
        dests.append("toast")
    if (notif.get("notify_me") or {}).get("access_code"):
        dests.append("alexa")
    print(f"Configured notification destinations: "
          f"{', '.join(dests) if dests else '(none)'}")
    print()

    overall_ok = True
    for account in cfg["accounts"]:
        name = account["name"]
        slug = account["_slug"]
        provider = make_provider(cfg, account)
        print("-" * 60)
        print(f"Account: {name}  [{account['provider']}]  (slug: {slug})")
        target = account.get("user_email") or account.get("email") or "?"
        print(f"  Mailbox: {target}")
        print(f"  Scan folders: {', '.join(account.get('scan_folders', []))}")
        ch = account.get("channels", {})
        print(f"  Account channels: "
              f"pushover={ch.get('pushover')} toast={ch.get('toast')} "
              f"alexa={ch.get('alexa')}")

        print("  Verifying auth/connection...")
        try:
            provider.connect()
            print("    OK")
        except Exception as e:
            print(f"    FAILED: {e}")
            overall_ok = False

        trigs = enabled_triggers(account)
        print(f"  Enabled triggers: {len(trigs)}")
        for t in trigs:
            eff = resolve_channels(account, t)
            on = [k for k, v in eff.items() if v]
            src = "trigger-override" if isinstance(
                t.get("channels"), dict) else "account-default"
            print(f"    - {t.get('name')}: "
                  f"domains={t.get('sender_domains', [])} "
                  f"addrs={t.get('sender_addresses', [])} "
                  f"subj_kw={t.get('subject_keywords', [])} "
                  f"body_kw={t.get('body_keywords', [])} "
                  f"channels={on} ({src})")
        print()

    print("=" * 60)
    if overall_ok:
        print("Test passed. Run without arguments to start the watchdog.")
        return 0
    print("Test FAILED for one or more accounts. Fix the errors above.")
    print("Microsoft accounts that failed auth may need "
          "'python watchdog.py --setup'.")
    return 1


def _mode_once(cfg):
    """--once: run a single poll pass across all accounts, then exit."""
    print("urgent-email-watch v2 -- single pass")
    providers = build_providers(cfg)
    total = run_all_accounts(cfg, providers, do_pending=True)
    print(f"Pass complete. {total} alert(s) fired across all accounts.")
    return 0


def main(argv):
    """Entry point. Dispatches to a mode based on the first argument."""
    arg = argv[1] if len(argv) > 1 else ""

    if arg in ("-h", "--help"):
        print(__doc__)
        return 0

    try:
        cfg = load_config()
    except ConfigError as e:
        print(f"CONFIG ERROR: {e}", file=sys.stderr)
        return 2

    # Enable file logging for all modes -- the installer runs headless.
    init_log_file(os.path.join(cfg["_state_dir_resolved"], "watchdog.log"))

    if arg == "--setup":
        return _mode_setup(cfg)
    if arg == "--test":
        return _mode_test(cfg)
    if arg == "--once":
        return _mode_once(cfg)
    if arg:
        print(f"Unknown argument: {arg}", file=sys.stderr)
        print("Valid modes: --setup, --test, --once, --help, or no argument "
              "(loop).", file=sys.stderr)
        return 2

    providers = build_providers(cfg)
    watchdog_loop(cfg, providers)
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
