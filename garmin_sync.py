"""
garmin_sync.py
Fetches recent Garmin Connect activities and syncs them to a dedicated
Google Calendar ('Garmin Workouts') as events. De-duplicates so it's safe
to run daily. Windows Task Scheduler-friendly.

First-run:
  - Prompts for Garmin login (credentials stored via garth tokens)
  - Opens browser for Google OAuth consent (token cached locally)
  - Creates the 'Garmin Workouts' calendar if missing

Subsequent runs: silent, no prompts.
"""

from __future__ import annotations

import os
import sys
import logging
from datetime import datetime, timedelta
from pathlib import Path

from dotenv import load_dotenv

# --- Garmin ---
from garminconnect import (
    Garmin,
    GarminConnectAuthenticationError,
    GarminConnectConnectionError,
    GarminConnectTooManyRequestsError,
)

try:
    import garth  # noqa: F401
    from garth.exc import GarthException, GarthHTTPError  # noqa: F401
except Exception:
    garth = None
    GarthException = Exception
    GarthHTTPError = Exception

# --- Google Calendar ---
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError


# -------------------- Config --------------------

SCRIPT_DIR = Path(__file__).resolve().parent
load_dotenv(SCRIPT_DIR / ".env")

GARMIN_EMAIL = os.getenv("GARMIN_EMAIL")
GARMIN_PASSWORD = os.getenv("GARMIN_PASSWORD")

GARMIN_TOKEN_DIR = SCRIPT_DIR / ".garminconnect"
GOOGLE_CLIENT_SECRETS = SCRIPT_DIR / "credentials.json"
GOOGLE_TOKEN_FILE = SCRIPT_DIR / "token.json"

CALENDAR_NAME = "Garmin Workouts"
SYNC_DAYS_BACK = 7
TIMEZONE = os.getenv("TIMEZONE", "America/New_York")

GOOGLE_SCOPES = ["https://www.googleapis.com/auth/calendar"]

LOG_FILE = SCRIPT_DIR / "garmin_sync.log"
logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    handlers=[
        logging.FileHandler(LOG_FILE, encoding="utf-8"),
        logging.StreamHandler(sys.stdout),
    ],
)
log = logging.getLogger("garmin_sync")


# -------------------- Garmin --------------------

def _save_tokens(g: Garmin) -> None:
    """Persist Garmin session tokens. Tolerates garminconnect version variation."""
    GARMIN_TOKEN_DIR.mkdir(parents=True, exist_ok=True)
    path = str(GARMIN_TOKEN_DIR)
    attempts = [
        lambda: g.garth.dump(path),
        lambda: garth.save(path) if garth else None,
        lambda: g.dump_session(path) if hasattr(g, "dump_session") else None,
    ]
    for attempt in attempts:
        try:
            attempt()
            log.info("Garmin: cached tokens to %s", path)
            return
        except AttributeError:
            continue
        except Exception as e:
            log.debug("Token save variant failed: %s", e)
            continue
    log.warning("Garmin: could not persist tokens — next run will re-login.")


def _is_rate_limited(err: Exception) -> bool:
    msg = str(err)
    return (
        "429" in msg
        or "Too Many Requests" in msg
        or isinstance(err, GarminConnectTooManyRequestsError)
    )


def garmin_login() -> Garmin:
    """Log in to Garmin, reusing cached tokens when possible."""
    # 1. Try resume from cached tokens.
    if GARMIN_TOKEN_DIR.exists() and any(GARMIN_TOKEN_DIR.iterdir()):
        try:
            g = Garmin()
            g.login(str(GARMIN_TOKEN_DIR))
            log.info("Garmin: reused cached tokens")
            return g
        except Exception as e:
            log.info("Garmin: cached tokens unusable (%s), doing fresh login", e)

    # 2. Fresh login requires credentials.
    if not GARMIN_EMAIL or not GARMIN_PASSWORD:
        raise RuntimeError(
            "GARMIN_EMAIL and GARMIN_PASSWORD must be set in .env for first-time login."
        )

    g = Garmin(email=GARMIN_EMAIL, password=GARMIN_PASSWORD)
    try:
        result = g.login()
    except Exception as e:
        if _is_rate_limited(e):
            raise RuntimeError(
                "Garmin rate-limited this IP (HTTP 429). "
                "Wait 1-24 hours, or try from another network (e.g. phone hotspot). "
                "Repeated failed logins or running too often triggers this."
            ) from e
        raise

    # 3. Handle MFA if required.
    if isinstance(result, tuple) and len(result) >= 1 and result[0] == "needs_mfa":
        try:
            code = input("Garmin MFA code (from email/authenticator): ").strip()
        except EOFError as e:
            raise RuntimeError(
                "MFA required but no interactive input available. "
                "Run the script manually the first time to complete MFA."
            ) from e
        if hasattr(g, "resume_login"):
            g.resume_login(result[1], code)
        else:
            raise RuntimeError(
                "Installed garminconnect version doesn't support resume_login. "
                "Run: pip install -U garminconnect"
            )

    _save_tokens(g)
    return g


def fetch_activities(g: Garmin, days_back: int):
    start = datetime.now().date() - timedelta(days=days_back)
    end = datetime.now().date()
    log.info("Garmin: fetching activities %s -> %s", start, end)
    acts = g.get_activities_by_date(start.isoformat(), end.isoformat())
    log.info("Garmin: got %d activities", len(acts))
    return acts


# -------------------- Google Calendar --------------------

def google_service():
    creds = None
    if GOOGLE_TOKEN_FILE.exists():
        creds = Credentials.from_authorized_user_file(str(GOOGLE_TOKEN_FILE), GOOGLE_SCOPES)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            if not GOOGLE_CLIENT_SECRETS.exists():
                raise RuntimeError(
                    f"Missing Google OAuth client file: {GOOGLE_CLIENT_SECRETS}. "
                    "Download it from Google Cloud Console (see README)."
                )
            flow = InstalledAppFlow.from_client_secrets_file(
                str(GOOGLE_CLIENT_SECRETS), GOOGLE_SCOPES
            )
            creds = flow.run_local_server(port=0)
        GOOGLE_TOKEN_FILE.write_text(creds.to_json(), encoding="utf-8")
    return build("calendar", "v3", credentials=creds, cache_discovery=False)


def get_or_create_calendar(service, name: str) -> str:
    page_token = None
    while True:
        cal_list = service.calendarList().list(pageToken=page_token).execute()
        for entry in cal_list.get("items", []):
            if entry.get("summary") == name:
                log.info("GCal: using existing calendar '%s' (%s)", name, entry["id"])
                return entry["id"]
        page_token = cal_list.get("nextPageToken")
        if not page_token:
            break
    created = service.calendars().insert(body={
        "summary": name,
        "description": "Auto-synced from Garmin Connect",
        "timeZone": TIMEZONE,
    }).execute()
    log.info("GCal: created calendar '%s' (%s)", name, created["id"])
    return created["id"]


def activity_to_event(act: dict) -> dict:
    activity_id = str(act.get("activityId"))
    name = act.get("activityName") or act.get("activityType", {}).get("typeKey", "Activity")
    type_key = act.get("activityType", {}).get("typeKey", "activity").replace("_", " ").title()

    start_local = act.get("startTimeLocal")
    duration_s = int(act.get("duration") or 0)

    start_dt = datetime.fromisoformat(start_local)
    end_dt = start_dt + timedelta(seconds=duration_s if duration_s > 0 else 60)

    distance_m = act.get("distance") or 0
    distance_km = distance_m / 1000.0
    distance_mi = distance_m / 1609.344
    calories = act.get("calories") or 0
    avg_hr = act.get("averageHR")
    max_hr = act.get("maxHR")
    avg_speed = act.get("averageSpeed")

    pace_str = ""
    if avg_speed and avg_speed > 0 and distance_m > 0:
        pace_s_per_mi = 1609.344 / avg_speed
        m, s = divmod(int(pace_s_per_mi), 60)
        pace_str = f"{m}:{s:02d} /mi"

    dur_h, rem = divmod(duration_s, 3600)
    dur_m, dur_s = divmod(rem, 60)
    dur_str = f"{dur_h}h {dur_m}m {dur_s}s" if dur_h else f"{dur_m}m {dur_s}s"

    desc_lines = [
        f"Type: {type_key}",
        f"Duration: {dur_str}",
        f"Distance: {distance_km:.2f} km ({distance_mi:.2f} mi)" if distance_m else "",
        f"Calories: {calories}" if calories else "",
        f"Avg HR: {avg_hr} bpm" if avg_hr else "",
        f"Max HR: {max_hr} bpm" if max_hr else "",
        f"Avg Pace: {pace_str}" if pace_str else "",
        "",
        f"https://connect.garmin.com/modern/activity/{activity_id}",
    ]
    description = "\n".join(line for line in desc_lines if line != "")

    summary = f"{type_key}: {name}"
    if distance_m:
        summary = f"{type_key} {distance_mi:.2f}mi"

    return {
        "summary": summary,
        "description": description,
        "start": {"dateTime": start_dt.isoformat(), "timeZone": TIMEZONE},
        "end": {"dateTime": end_dt.isoformat(), "timeZone": TIMEZONE},
        # Stable ID for dedup. Google requires [a-v0-9], 5-1024 chars.
        "id": f"garmin{activity_id}",
    }


def upsert_event(service, calendar_id: str, event_body: dict) -> str:
    event_id = event_body["id"]
    try:
        service.events().get(calendarId=calendar_id, eventId=event_id).execute()
        service.events().update(
            calendarId=calendar_id, eventId=event_id, body=event_body
        ).execute()
        return "updated"
    except HttpError as e:
        if e.resp.status == 404:
            service.events().insert(calendarId=calendar_id, body=event_body).execute()
            return "created"
        raise


# -------------------- Main --------------------

def main() -> int:
    log.info("=" * 50)
    log.info("Garmin -> Google Calendar sync starting")
    try:
        g = garmin_login()
        activities = fetch_activities(g, SYNC_DAYS_BACK)

        service = google_service()
        cal_id = get_or_create_calendar(service, CALENDAR_NAME)

        created = updated = skipped = 0
        for act in activities:
            try:
                body = activity_to_event(act)
                result = upsert_event(service, cal_id, body)
                if result == "created":
                    created += 1
                elif result == "updated":
                    updated += 1
                log.info("  %s: %s", result, body["summary"])
            except Exception as e:
                skipped += 1
                log.warning("  skipped activity %s: %s", act.get("activityId"), e)

        log.info("Done. created=%d updated=%d skipped=%d", created, updated, skipped)
        return 0
    except Exception as e:
        log.exception("Sync failed: %s", e)
        return 1


if __name__ == "__main__":
    sys.exit(main())
