-- Daily Truck Tracker — projected-hours store (Cloudflare D1)
-- One row per truck per day. Team enters proj_hrs + notes; the rest is a
-- snapshot captured at entry time so the proj-vs-actual record stays stable
-- even if the underlying dispatch/Payworks data shifts later.
CREATE TABLE IF NOT EXISTS proj_hours (
  date        TEXT    NOT NULL,           -- YYYY-MM-DD (operational day)
  truck       TEXT    NOT NULL,           -- tracker truck id (tnum)
  proj_hrs    REAL,                       -- team-entered projected labour hours
  notes       TEXT,                       -- team-entered notes
  driver      TEXT,                       -- snapshot
  bu          TEXT,                       -- snapshot (GTA / East / West / GTA Tents)
  revenue     REAL,                       -- snapshot (truck revenue that day)
  stops       INTEGER,                    -- snapshot
  actual_hrs  REAL,                       -- Payworks actual (filled when known)
  entered_by  TEXT,                       -- Cloudflare Access email, if available
  entered_at  TEXT,                       -- ISO timestamp of first entry
  updated_at  TEXT,                       -- ISO timestamp of last edit
  PRIMARY KEY (date, truck)
);
CREATE INDEX IF NOT EXISTS idx_proj_hours_date ON proj_hours(date);
