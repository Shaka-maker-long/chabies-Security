# Chabies Security / Studio Delta

The public website is at the repo root. **Studio Delta Production** (shop floor + office Orders and Enquiries) lives in [`studio-delta-production/`](studio-delta-production/).

Floor start, pause, resume, finish, durations, steel usage, and backboard usage save on Railway in `/app/data` as **SQLite** (`studio-delta.db`), in the same `orders` and `users` tables as `/orders` and Users. Google Sheets is only a one-time import. Office **Enquiries** is on that same SQLite file.

See `studio-delta-production/README.md` for Railway setup.
