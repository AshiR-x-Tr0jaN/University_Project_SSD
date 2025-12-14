
-- Scans table
CREATE TABLE IF NOT EXISTS scans (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    target_url TEXT NOT NULL,
    scan_type TEXT NOT NULL,
    start_time TEXT NOT NULL,
    end_time TEXT,
    total_alerts INTEGER DEFAULT 0,
    high_risk INTEGER DEFAULT 0,
    medium_risk INTEGER DEFAULT 0,
    low_risk INTEGER DEFAULT 0,
    status TEXT DEFAULT 'running',
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
);

-- Vulnerabilities table
CREATE TABLE IF NOT EXISTS vulnerabilities (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    scan_id INTEGER NOT NULL,
    alert_name TEXT NOT NULL,
    risk_level TEXT NOT NULL,
    confidence TEXT,
    url TEXT,
    description TEXT,
    solution TEXT,
    reference TEXT,
    cwe_id INTEGER,
    wasc_id INTEGER,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
    FOREIGN KEY (scan_id) REFERENCES scans(id) ON DELETE CASCADE
);

-- Index for faster queries
CREATE INDEX IF NOT EXISTS idx_scan_id ON vulnerabilities(scan_id);
CREATE INDEX IF NOT EXISTS idx_risk_level ON vulnerabilities(risk_level);
CREATE INDEX IF NOT EXISTS idx_target_url ON scans(target_url);
