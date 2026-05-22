-- ============================================================
-- CPE Refurb Manager v2.0 — Supabase Schema
-- Run this entire file in Supabase → SQL Editor → New query
-- ============================================================

-- ── DEVICES ──────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS devices (
  id                TEXT        PRIMARY KEY,
  serial            TEXT        NOT NULL UNIQUE,
  mac               TEXT        DEFAULT '',
  model             TEXT        DEFAULT '',
  type              TEXT        NOT NULL
                                CHECK (type IN ('Router/Modem','Set-top Box','ONT/OLT')),
  stage             TEXT        NOT NULL DEFAULT 'Triage'
                                CHECK (stage IN (
                                  'Intake','Triage','Refurbishment','QC Check',
                                  'Stock','Scrap','Escalated','ECUS',
                                  'In Transit','Confirmed'
                                )),
  outcome           TEXT        CHECK (outcome IN ('Working','Not Working','Scrap')),
  received_date     DATE        NOT NULL DEFAULT CURRENT_DATE,
  notes             TEXT        DEFAULT '',
  sent_to_partner   BOOLEAN     DEFAULT FALSE,
  partner_outcome   TEXT        CHECK (partner_outcome IN ('Working','Not Working')),
  partner_notes     TEXT        DEFAULT '',
  pending_action    TEXT        CHECK (pending_action IN ('Refurbishment','Scrap','ECUS')),
  created_at        TIMESTAMPTZ DEFAULT NOW(),
  updated_at        TIMESTAMPTZ DEFAULT NOW()
);

-- Auto-update updated_at on any row change
CREATE OR REPLACE FUNCTION update_updated_at()
RETURNS TRIGGER AS $$
BEGIN NEW.updated_at = NOW(); RETURN NEW; END;
$$ LANGUAGE plpgsql;

DROP TRIGGER IF EXISTS trg_devices_updated_at ON devices;
CREATE TRIGGER trg_devices_updated_at
  BEFORE UPDATE ON devices
  FOR EACH ROW EXECUTE FUNCTION update_updated_at();

-- ── USERS ─────────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS users (
  id          UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  name        TEXT        NOT NULL,
  username    TEXT        NOT NULL UNIQUE,
  password    TEXT        NOT NULL,
  role        TEXT        NOT NULL DEFAULT 'stock'
              CHECK (role IN ('admin','stock','partner')),
  created_at  TIMESTAMPTZ DEFAULT NOW()
);

-- ── UPLOAD LOGS ───────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS upload_logs (
  id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  log_type     TEXT        NOT NULL CHECK (log_type IN ('intake','partner')),
  total_rows   INTEGER     DEFAULT 0,
  added        INTEGER     DEFAULT 0,
  skipped      INTEGER     DEFAULT 0,
  matched      INTEGER     DEFAULT 0,
  unmatched    INTEGER     DEFAULT 0,
  invalid_rows INTEGER     DEFAULT 0,
  created_at   TIMESTAMPTZ DEFAULT NOW()
);

-- ── AUDIT LOG ─────────────────────────────────────────────────
CREATE TABLE IF NOT EXISTS audit_log (
  id           UUID        PRIMARY KEY DEFAULT gen_random_uuid(),
  device_id    TEXT        REFERENCES devices(id) ON DELETE CASCADE,
  action       TEXT        NOT NULL,
  old_stage    TEXT,
  new_stage    TEXT,
  old_outcome  TEXT,
  new_outcome  TEXT,
  performed_by TEXT        DEFAULT 'system',
  created_at   TIMESTAMPTZ DEFAULT NOW()
);

-- ── INDEXES ───────────────────────────────────────────────────
CREATE INDEX IF NOT EXISTS idx_devices_stage        ON devices(stage);
CREATE INDEX IF NOT EXISTS idx_devices_type         ON devices(type);
CREATE INDEX IF NOT EXISTS idx_devices_outcome      ON devices(outcome);
CREATE INDEX IF NOT EXISTS idx_devices_partner      ON devices(sent_to_partner);
CREATE INDEX IF NOT EXISTS idx_devices_serial       ON devices(serial);
CREATE INDEX IF NOT EXISTS idx_audit_device         ON audit_log(device_id);
CREATE INDEX IF NOT EXISTS idx_upload_logs_type     ON upload_logs(log_type);

-- ── DISABLE RLS (prototype) ───────────────────────────────────
-- Enable and add row-level policies before going to production
ALTER TABLE devices     DISABLE ROW LEVEL SECURITY;
ALTER TABLE users       DISABLE ROW LEVEL SECURITY;
ALTER TABLE upload_logs DISABLE ROW LEVEL SECURITY;
ALTER TABLE audit_log   DISABLE ROW LEVEL SECURITY;

-- ── SEED DEMO DATA ────────────────────────────────────────────
-- Safe to re-run: deletes existing seed rows first
DELETE FROM devices WHERE id LIKE 'D-00%';
DELETE FROM users   WHERE username IN ('admin','stock','partner');

-- Default users
INSERT INTO users (name, username, password, role) VALUES
  ('Administrator',   'admin',   'admin123',   'admin'),
  ('Stock Manager',   'stock',   'stock123',   'stock'),
  ('Refurb Partner',  'partner', 'partner123', 'partner')
ON CONFLICT (username) DO NOTHING;

-- Seed devices across all stages
INSERT INTO devices (id, serial, mac, model, type, stage, outcome, received_date,
                     notes, sent_to_partner, partner_outcome, partner_notes, pending_action)
VALUES
  -- Ready (Stock)
  ('D-0001','SN-88421','E8:65:D4:11:22:33','Huawei HG8245H',       'Router/Modem','Stock',        'Working',     '2025-05-01','Firmware reset',      TRUE,  NULL, '',   NULL),
  ('D-0002','SN-33190','28:C6:8E:44:55:66','Technicolor TC7200',    'Set-top Box', 'Stock',        'Working',     '2025-05-02','HDMI port replaced',  TRUE,  NULL, '',   NULL),
  ('D-0003','SN-77042','9C:97:26:77:88:99','Nokia G-010G-P',        'ONT/OLT',     'Stock',        'Working',     '2025-05-03','Clean & reconfigure', TRUE,  NULL, '',   NULL),
  ('D-0010','SN-47832','00:1A:2B:33:44:55','Cisco DPC3825',         'Router/Modem','Stock',        'Working',     '2025-05-10','',                    TRUE,  NULL, '',   NULL),
  -- Scrap
  ('D-0004','SN-55610','34:4B:50:AA:BB:CC','ZTE ZXHN H108N',        'Router/Modem','Scrap',        'Scrap',       '2025-05-04','PCB damage',          FALSE, NULL, '',   NULL),
  ('D-0005','SN-12983','',                 '',                      'Set-top Box', 'Scrap',        'Scrap',       '2025-05-05','Burned PSU',          FALSE, NULL, '',   NULL),
  -- Refurbishment queue (not yet dispatched)
  ('D-0006','SN-66101','54:89:98:DD:EE:FF','Huawei EchoLife EG8145','ONT/OLT',     'Refurbishment',NULL,          '2025-05-06','In repair',           FALSE, NULL, '',   NULL),
  ('D-0007','SN-29344','C8:D7:19:12:34:56','Sagemcom F@ST 5366',    'Router/Modem','Refurbishment',NULL,          '2025-05-07','Awaiting dispatch',   FALSE, NULL, '',   NULL),
  -- Triage queue
  ('D-0008','SN-84720','',                 '',                      'Set-top Box', 'Triage',       NULL,          '2025-05-08','',                    FALSE, NULL, '',   NULL),
  -- Escalated
  ('D-0009','SN-39011','BC:F6:12:AB:CD:EF','ZTE F660',              'ONT/OLT',     'Escalated',    'Not Working', '2025-05-09','Failed QC twice',     TRUE,  NULL, '',   NULL),
  -- Refurbishment (dispatched, awaiting partner)
  ('D-0011','SN-90012','7C:4C:A5:66:77:88','Sagemcom FAST 3686',    'Set-top Box', 'Refurbishment',NULL,          '2025-05-11','Screen replaced',     TRUE,  NULL, '',   NULL),
  -- ECUS
  ('D-0012','SN-55123','48:57:54:43:CB:5B','Cisco SB6141',          'Router/Modem','ECUS',         NULL,          '2025-05-12','Pending review',      FALSE, NULL, '',   NULL),
  -- Confirmed (awaiting return to partner)
  ('D-0013','SN-21001','E8:65:D4:AA:BB:CC','Huawei HG8245Q',        'Router/Modem','Confirmed',    'Working',     '2025-05-13','',                    TRUE,  NULL, '',   NULL),
  -- In Transit
  ('D-0014','SN-32002','9C:97:26:11:22:33','Nokia G-010S-A',        'ONT/OLT',     'In Transit',   'Not Working', '2025-05-14','Faulty power supply', TRUE,  NULL, '',   NULL)
ON CONFLICT (id) DO NOTHING;
