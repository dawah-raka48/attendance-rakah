# ZKTeco device sync agent (Windows)

This agent reads user IDs and raw attendance events from the ZKTeco MB1000/ID over the local network and sends them to the Attendance Rakah Apps Script endpoint.

## Safety
- Read-only device access: no enrollments, settings changes, or deletions.
- Raw device status and punch values are preserved. The agent does not guess check-in/check-out meaning or shift.
- Duplicate events are ignored by a deterministic event fingerprint.
- The sync key is entered on first run and saved locally as `config.json`. Never commit or share that file.

## Setup
1. In Apps Script, replace the backend with the repository's `backend/Code.gs`.
2. Run `setupAttendanceDatabase()` once from the Apps Script editor to create the `deviceLogs` sheet. Existing sheets and rows are preserved.
3. In Apps Script Project Settings > Script Properties, add:
   - Property: `DEVICE_SYNC_KEY`
   - Value: a long random secret created by the administrator.
4. Update the existing Web App deployment to a new version.
5. Download `agent.py` to a private folder on the Windows computer that can reach the device.
6. Run `python agent.py`. On first run, paste the sync key when prompted.
7. Keep the computer on; the agent polls every 60 seconds. Press Ctrl+C to stop.

Device defaults in this agent: `192.168.100.76:4370`. Change these constants only if the device address changes.

## Destination
Raw events are stored in the Google Sheet tab `deviceLogs`. Employee records are created with temporary names (e.g. `موظف 4`) and department `غير محدد` if their device ID is not already in the employees table.

These raw logs are intentionally separate from attendance reports until the device's status/punch codes are verified against actual arrival/departure behavior.
