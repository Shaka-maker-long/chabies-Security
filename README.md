# Chabies Security / Studio Delta

The public website is at the repo root. **Studio Delta Production** (shop floor + office Orders and Enquiries) lives in [`studio-delta-production/`](studio-delta-production/).

Floor start, pause, resume, finish, durations, steel usage, and backboard usage save on Railway in `/app/data`, in the same `ORDERS` table as `/orders`. Google Sheets is only a one-time import. Office **Enquiries** is a Google-Sheet-style log on that same Railway store.

See `studio-delta-production/README.md` for Railway setup.
