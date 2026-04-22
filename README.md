# Garmin -> Google Calendar Sync

Daily sync of your Garmin Connect activities into a dedicated "Garmin Workouts" Google Calendar. De-duplicates on re-run.

## What you get
- `garmin_sync.py` — the script
- `requirements.txt` — Python deps
- `.env.example` — credentials template
- `run_garmin_sync.bat` — Windows Task Scheduler wrapper
- Logs to `garmin_sync.log` next to the script

## One-time setup

### 1. Install Python deps
Put everything in one folder (e.g. `C:\Tools\garmin-sync\`). From that folder:

```
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
```

### 2. Garmin credentials
Copy `.env.example` to `.env` and fill in your Garmin email/password. Used only on first login; after that, token cache keeps you signed in.

### 3. Google Calendar API — OAuth client
1. Go to https://console.cloud.google.com/
2. Create a new project (or reuse one).
3. **APIs & Services → Library** → enable **Google Calendar API**.
4. **APIs & Services → OAuth consent screen** → choose **External**, fill in app name + your email, add yourself as a test user. You can skip all scopes on this screen.
5. **APIs & Services → Credentials → Create Credentials → OAuth client ID** → app type **Desktop app**.
6. Download the JSON. Rename it to `credentials.json` and drop it in the script folder.

### 4. First run (interactive)
```
.venv\Scripts\python.exe garmin_sync.py
```
- It will log into Garmin (may prompt for MFA code if your account has 2FA).
- A browser window opens for Google consent → approve.
- Token files (`token.json`, `.garminconnect/`) are saved for silent future runs.

## Schedule daily on Windows

1. Open **Task Scheduler**.
2. **Create Basic Task** → name it "Garmin Calendar Sync".
3. Trigger: **Daily**, pick a time (e.g. 11:00 PM so the day's workouts are in).
4. Action: **Start a program**
   - Program/script: `run_garmin_sync.bat`
   - Start in: the folder you put the files in (e.g. `C:\Tools\garmin-sync\`)
5. Finish. Right-click the task → **Properties** → check **Run whether user is logged on or not** if you want it to run headless.

## Notes

- **Dedup**: each Garmin activity gets a stable Google event ID (`garmin<activityId>`), so re-running is safe.
- **Window**: looks back 7 days each run to catch late-arriving uploads.
- **Calendar**: creates "Garmin Workouts" automatically on first run.
- **2FA**: if your Garmin account has MFA, the first login will prompt in the terminal. Run manually the first time.
- **Timezone**: default `America/New_York`. Override with `TIMEZONE=` in `.env`.

## Troubleshooting

- `Missing Google OAuth client file`: download `credentials.json` (step 3 above) into the script folder.
- `GARMIN_EMAIL and GARMIN_PASSWORD must be set`: create `.env` from `.env.example`.
- Stale Garmin session: delete the `.garminconnect/` folder and re-run.
- Stale Google auth: delete `token.json` and re-run.
- Check `garmin_sync.log` for details.
