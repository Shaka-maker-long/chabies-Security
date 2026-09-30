# Studio Delta security

How this app maps to the shop security checklist. Archive a copy of this file with your off-site backups so the rules stay on your system as well as in the repo.

| # | Area | What we ship | Status |
|---|---|---|---|
| 1 | Login | Hashed access codes (scrypt), login throttling, **no first-login auto-approve**. Clear on-screen messages for incorrect name/access code vs device waiting for Manager approval. Every device must be Manager-approved and assigned to that person. | Critical — done |
| 2 | Permissions | Role-based access: Manager / Admin / Marketing / Production. Least privilege on office APIs (`canManageUsers`, debtors, marketing write allow-list). | Critical — done |
| 3 | Database | SQLite on the private Railway volume (`DATA_DIR`), not a public Postgres URL. Restrict who can download `.db` / `.tgz` (Manager only). | Critical — done (volume-private; use Railway private networking) |
| 4 | Encryption | HTTPS/TLS via Railway. Access codes hashed at rest. Session cookies `HttpOnly` + `SameSite=Lax` + `Secure` on HTTPS. Full disk / volume encryption is on the host (Railway). | Critical — transit + code hashing done; volume at-rest follows host |
| 5 | Audit trail | `DATA_DIR/audit-log.jsonl` — logins, failures, throttles, device approve/revoke, backup downloads. Manager reads via `GET /api/office/audit`. | Critical — done |
| 6 | Backups | Automated `.tgz` on the volume + optional Google Drive archive (`BACKUP_DRIVE_FOLDER_ID`). Manager-only download/restore. Download archives to your PC/Drive regularly. | Critical — done |
| 7 | Sessions | Random tokens, 14-day expiry, logout clears cookie, revoke device drops bound sessions, sessions store `deviceId`. | Critical — done |
| 8 | API security | Every `/api/office/*` route (except login) requires session + role checks. `/api/run` requires a session except login / user list. Device gate on login. | Critical — done |
| 9 | Secrets | Env vars only (`.env.example`). Never commit passwords or PATs. | Critical — done |
| 10 | Monitoring | Audit log + failed-login throttle events. Review `/api/office/audit` and Users → Devices pending list. | High — foundation done |
| 11 | Updates | Keep Node 22+, `npm audit` / dependency bumps on a schedule. Patch Railway stack. | High — process |
| 12 | Recovery | Users → Backup restore (RESTORE + Manager code twice). Keep Drive / local archives. | High — runbooks in README |

## Device approval (no auto-trust)

1. Manager logs in on an already-approved device.
2. When someone (or the Manager on a new phone) tries a device that is not linked yet, login shows **waiting for Manager approval**.
3. Manager opens **Users → Devices**, assigns the phone to that person, and approves it.
4. Revoke a lost phone → their sessions end immediately.

Optional emergency only: `DEVICE_BOOTSTRAP_CODE` on the server can still unlock a Manager device via API if every Manager phone is lost — it is not shown on the login screen.

## Archive on your system

- Download **Users → Backup** complete `.tgz` to your PC or shared Drive folder.
- Keep this `SECURITY.md` next to those archives.
- After any restore, re-check pending devices and Manager job title.

## MFA (next step)

App-level TOTP / email one-time codes are not shipped yet. Prefer device approval + hashed codes + throttle now; add MFA when you are ready for authenticator apps on every phone.
