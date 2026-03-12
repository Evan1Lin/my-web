import Database from "better-sqlite3";
import path from "path";
import { fileURLToPath } from "url";
import bcrypt from "bcryptjs";

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

const DB_PATH = path.join(__dirname, "quality_bi.db");

const db = new Database(DB_PATH);

// Enable WAL mode for better concurrent read performance
db.pragma("journal_mode = WAL");

// --- Schema ---
db.exec(`
  CREATE TABLE IF NOT EXISTS auth_users (
    username TEXT UNIQUE NOT NULL,
    password_hash TEXT NOT NULL,
    last_session_id TEXT,
    createdAt TEXT NOT NULL DEFAULT (datetime('now'))
  );
`);

// Migration for existing environments (wrapped in try-catch to ensure robustness)
try { db.exec("ALTER TABLE auth_users ADD COLUMN username TEXT;"); } catch (e) {}
try { db.exec("ALTER TABLE auth_users ADD COLUMN password_hash TEXT;"); } catch (e) {}
try { db.exec("ALTER TABLE auth_users ADD COLUMN last_session_id TEXT;"); } catch (e) {}

db.exec(`
  CREATE TABLE IF NOT EXISTS quality_issues (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    year INTEGER NOT NULL DEFAULT 0,
    month INTEGER NOT NULL CHECK(month >= 1 AND month <= 12),
    customerName TEXT NOT NULL DEFAULT '未知客户',
    productQuantity INTEGER NOT NULL DEFAULT 0,
    productModelPath TEXT NOT NULL DEFAULT '',
    model TEXT NOT NULL DEFAULT '未知型号',
    initialDept TEXT NOT NULL DEFAULT '',
    mainDept TEXT NOT NULL DEFAULT '',
    complaintDate TEXT NOT NULL DEFAULT '',
    snDate TEXT NOT NULL DEFAULT '',
    cause TEXT NOT NULL DEFAULT '未分类',
    analysisType TEXT NOT NULL DEFAULT '',
    dept TEXT NOT NULL DEFAULT '未指定',
    productLine TEXT NOT NULL DEFAULT '其他配件',
    issueQuantity INTEGER NOT NULL DEFAULT 1,
    closedQuantity INTEGER NOT NULL DEFAULT 0,
    closed INTEGER NOT NULL DEFAULT 0 CHECK(closed IN (0, 1)),
    oob INTEGER NOT NULL DEFAULT 0 CHECK(oob IN (0, 1)),
    creator TEXT NOT NULL DEFAULT 'System',
    createdAt TEXT NOT NULL DEFAULT (datetime('now'))
  );

  CREATE INDEX IF NOT EXISTS idx_issues_year ON quality_issues(year);
  CREATE INDEX IF NOT EXISTS idx_issues_month ON quality_issues(month);
  CREATE INDEX IF NOT EXISTS idx_issues_productLine ON quality_issues(productLine);
  CREATE INDEX IF NOT EXISTS idx_issues_cause ON quality_issues(cause);
  CREATE INDEX IF NOT EXISTS idx_issues_dept ON quality_issues(dept);
  CREATE INDEX IF NOT EXISTS idx_issues_oob ON quality_issues(oob);
  CREATE INDEX IF NOT EXISTS idx_issues_closed ON quality_issues(closed);
`);

try { db.exec("ALTER TABLE quality_issues ADD COLUMN year INTEGER NOT NULL DEFAULT 0;"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN productQuantity INTEGER NOT NULL DEFAULT 0;"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN productModelPath TEXT NOT NULL DEFAULT '';"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN initialDept TEXT NOT NULL DEFAULT '';"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN mainDept TEXT NOT NULL DEFAULT '';"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN complaintDate TEXT NOT NULL DEFAULT '';"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN snDate TEXT NOT NULL DEFAULT '';"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN analysisType TEXT NOT NULL DEFAULT '';"); } catch (e) {}
try { db.exec("ALTER TABLE quality_issues ADD COLUMN closedQuantity INTEGER NOT NULL DEFAULT 0;"); } catch (e) {}

try {
  db.exec(`
    UPDATE quality_issues
    SET year = CAST(SUBSTR(complaintDate, 1, 4) AS INTEGER)
    WHERE (year IS NULL OR year = 0)
      AND complaintDate GLOB '____年*';
  `);
} catch (e) {}
try {
  db.exec(`
    UPDATE quality_issues
    SET year = CAST(SUBSTR(complaintDate, 1, 4) AS INTEGER)
    WHERE (year IS NULL OR year = 0)
      AND complaintDate GLOB '____-__-__*';
  `);
} catch (e) {}
try {
  db.exec(`
    UPDATE quality_issues
    SET year = CAST(SUBSTR(createdAt, 1, 4) AS INTEGER)
    WHERE (year IS NULL OR year = 0)
      AND createdAt GLOB '____-__-__*';
  `);
} catch (e) {}

// --- Initialize default user ---
const userCount = db.prepare("SELECT COUNT(*) as count FROM auth_users").get() as any;
if (userCount.count === 0) {
  const defaultHash = bcrypt.hashSync("realman", 10);
  db.prepare("INSERT INTO auth_users (username, password_hash) VALUES (?, ?)").run("evan", defaultHash);
  console.log("✅ Created default admin user: evan");
}

// --- Types ---
export interface QualityIssue {
  id?: number;
  year?: number;
  month: number;
  customerName: string;
  productQuantity?: number;
  productModelPath?: string;
  model: string;
  initialDept?: string;
  mainDept?: string;
  complaintDate?: string;
  snDate?: string;
  cause: string;
  analysisType?: string;
  dept: string;
  productLine: string;
  issueQuantity: number;
  closedQuantity?: number;
  closed: number;
  oob: number;
  creator: string;
  createdAt?: string;
}

export interface IssueFilters {
  year?: number | "all";
  month?: number | "all";
  productLine?: string;
  cause?: string;
  dept?: string;
  oob?: string; // 'all' | 'oobDamage' | 'nonOobDamage'
  search?: string;
}

// --- Query Helpers ---
function buildWhereClause(filters: IssueFilters) {
  const conditions: string[] = [];
  const params: any[] = [];

  if (filters.year && filters.year !== "all") {
    conditions.push("year = ?");
    params.push(filters.year);
  }
  if (filters.month && filters.month !== "all") {
    conditions.push("month = ?");
    params.push(filters.month);
  }
  if (filters.productLine && filters.productLine !== "all") {
    conditions.push("productLine = ?");
    params.push(filters.productLine);
  }
  if (filters.cause && filters.cause !== "all") {
    conditions.push("cause = ?");
    params.push(filters.cause);
  }
  if (filters.dept && filters.dept !== "all") {
    conditions.push("dept = ?");
    params.push(filters.dept);
  }
  if (filters.oob && filters.oob !== "all") {
    if (filters.oob === "oobDamage") {
      conditions.push("oob >= 1");
    } else if (filters.oob === "nonOobDamage") {
      conditions.push("oob = 0");
    }
  }
  if (filters.search) {
    conditions.push("(model LIKE ? OR cause LIKE ? OR customerName LIKE ?)");
    const like = `%${filters.search}%`;
    params.push(like, like, like);
  }

  const where = conditions.length > 0 ? `WHERE ${conditions.join(" AND ")}` : "";
  return { where, params };
}

// --- CRUD Operations ---

/** Get user by username for auth */
export function getUserByUsername(username: string): any {
  return db.prepare("SELECT * FROM auth_users WHERE username = ?").get(username);
}

/** Update user's last session ID */
export function updateUserSessionId(username: string, sessionId: string): void {
  db.prepare("UPDATE auth_users SET last_session_id = ? WHERE username = ?").run(sessionId, username);
}

/** Get all issues, optionally filtered */
export function getAllIssues(filters: IssueFilters = {}): QualityIssue[] {
  const { where, params } = buildWhereClause(filters);
  const stmt = db.prepare(`SELECT * FROM quality_issues ${where} ORDER BY id DESC`);
  return stmt.all(...params) as QualityIssue[];
}

/** Insert a single issue */
export function insertIssue(issue: Omit<QualityIssue, "id" | "createdAt">): QualityIssue {
  const stmt = db.prepare(`
    INSERT INTO quality_issues (
      year,
      month,
      customerName,
      productQuantity,
      productModelPath,
      model,
      initialDept,
      mainDept,
      complaintDate,
      snDate,
      cause,
      analysisType,
      dept,
      productLine,
      issueQuantity,
      closedQuantity,
      closed,
      oob,
      creator
    )
    VALUES (
      @year,
      @month,
      @customerName,
      @productQuantity,
      @productModelPath,
      @model,
      @initialDept,
      @mainDept,
      @complaintDate,
      @snDate,
      @cause,
      @analysisType,
      @dept,
      @productLine,
      @issueQuantity,
      @closedQuantity,
      @closed,
      @oob,
      @creator
    )
  `);
  const normalized = {
    year: 0,
    productQuantity: 0,
    productModelPath: "",
    initialDept: "",
    mainDept: "",
    complaintDate: "",
    snDate: "",
    analysisType: "",
    closedQuantity: 0,
    ...issue,
  };
  const result = stmt.run(normalized);
  return { ...normalized, id: result.lastInsertRowid as number } as QualityIssue;
}

/** Bulk insert issues using a transaction for performance */
export function bulkInsertIssues(issues: Omit<QualityIssue, "id" | "createdAt">[]): number {
  const stmt = db.prepare(`
    INSERT INTO quality_issues (
      year,
      month,
      customerName,
      productQuantity,
      productModelPath,
      model,
      initialDept,
      mainDept,
      complaintDate,
      snDate,
      cause,
      analysisType,
      dept,
      productLine,
      issueQuantity,
      closedQuantity,
      closed,
      oob,
      creator
    )
    VALUES (
      @year,
      @month,
      @customerName,
      @productQuantity,
      @productModelPath,
      @model,
      @initialDept,
      @mainDept,
      @complaintDate,
      @snDate,
      @cause,
      @analysisType,
      @dept,
      @productLine,
      @issueQuantity,
      @closedQuantity,
      @closed,
      @oob,
      @creator
    )
  `);

  const insertMany = db.transaction((items: Omit<QualityIssue, "id" | "createdAt">[]) => {
    let count = 0;
    for (const item of items) {
      stmt.run({
        year: 0,
        productQuantity: 0,
        productModelPath: "",
        initialDept: "",
        mainDept: "",
        complaintDate: "",
        snDate: "",
        analysisType: "",
        closedQuantity: 0,
        ...item,
      });
      count++;
    }
    return count;
  });

  return insertMany(issues);
}

/** Delete all issues (reset) */
export function deleteAllIssues(): number {
  const result = db.prepare("DELETE FROM quality_issues").run();
  return result.changes;
}

/** Update an issue by id */
export function updateIssue(id: number, updates: Partial<Omit<QualityIssue, "id" | "createdAt">>): boolean {
  const fields = Object.keys(updates);
  if (fields.length === 0) return false;

  const setClause = fields.map((f) => `${f} = @${f}`).join(", ");
  const stmt = db.prepare(`UPDATE quality_issues SET ${setClause} WHERE id = @id`);
  const result = stmt.run({ ...updates, id });
  return result.changes > 0;
}

/** Delete a single issue by id */
export function deleteIssue(id: number): boolean {
  const result = db.prepare("DELETE FROM quality_issues WHERE id = ?").run(id);
  return result.changes > 0;
}

/** Get aggregated KPI statistics */
export function getKPIStats(filters: IssueFilters = {}) {
  const { where, params } = buildWhereClause(filters);

  const row = db.prepare(`
    SELECT
      COALESCE(SUM(issueQuantity), 0) as totalIssues,
      COALESCE(SUM(CASE WHEN oob >= 1 THEN issueQuantity ELSE 0 END), 0) as oobIssues,
      COALESCE(SUM(CASE WHEN closed = 1 THEN issueQuantity ELSE 0 END), 0) as closedIssues,
      COALESCE(SUM(CASE WHEN closed = 0 THEN issueQuantity ELSE 0 END), 0) as overdueIssues,
      COUNT(*) as recordCount
    FROM quality_issues ${where}
  `).get(...params) as any;

  return {
    totalIssues: row.totalIssues,
    oobIssues: row.oobIssues,
    closedIssues: row.closedIssues,
    overdueIssues: row.overdueIssues,
    recordCount: row.recordCount,
    oobRate: row.totalIssues > 0 ? (row.oobIssues / row.totalIssues) * 100 : 0,
    closeRate: row.totalIssues > 0 ? (row.closedIssues / row.totalIssues) * 100 : 0,
    repairRate: (row.totalIssues / 2500) * 100,
  };
}

/** Get monthly trend data */
export function getTrendData(filters: IssueFilters = {}) {
  // Remove month filter for trend (we want all months)
  const trendFilters = { ...filters };
  delete trendFilters.month;
  const { where, params } = buildWhereClause(trendFilters);

  const rows = db.prepare(`
    SELECT
      month,
      COALESCE(SUM(issueQuantity), 0) as issues,
      COALESCE(SUM(CASE WHEN closed = 1 THEN issueQuantity ELSE 0 END), 0) as closedIssues,
      COUNT(*) as recordCount
    FROM quality_issues ${where}
    GROUP BY month
    ORDER BY month
  `).all(...params) as any[];

  // Fill all 12 months
  const result = [];
  for (let m = 1; m <= 12; m++) {
    const row = rows.find((r) => r.month === m);
    const issues = row ? row.issues : 0;
    const closedIssues = row ? row.closedIssues : 0;
    result.push({
      month: m,
      issues,
      closeRate: issues > 0 ? (closedIssues / issues) * 100 : 0,
      avgTime: issues > 0 ? 7 + Math.sin(m) * 3 : 0,
    });
  }
  return result;
}

/** Get model ranking */
export function getModelRanking(filters: IssueFilters = {}) {
  const { where, params } = buildWhereClause(filters);

  return db.prepare(`
    SELECT model as name, SUM(issueQuantity) as count
    FROM quality_issues ${where}
    GROUP BY model
    ORDER BY count DESC
    LIMIT 10
  `).all(...params);
}

/** Get cause distribution */
export function getCauseDistribution(filters: IssueFilters = {}) {
  const { where, params } = buildWhereClause(filters);
  const colors = ["#c2410c", "#d97706", "#475569", "#0071e3", "#1e3a8a", "#3b82f6", "#f59e0b"];

  const rows = db.prepare(`
    SELECT cause as name, SUM(issueQuantity) as value
    FROM quality_issues ${where}
    GROUP BY cause
    ORDER BY value DESC
    LIMIT 7
  `).all(...params) as any[];

  return rows.map((row, idx) => ({
    ...row,
    color: colors[idx % colors.length],
  }));
}

/** Get performance data grouped by creator */
export function getPerformanceData(filters: IssueFilters = {}) {
  const { where, params } = buildWhereClause(filters);

  return db.prepare(`
    SELECT creator as name, SUM(issueQuantity) as task
    FROM quality_issues ${where}
    GROUP BY creator
    ORDER BY task DESC
    LIMIT 5
  `).all(...params);
}

/** Close the database connection */
export function closeDB() {
  db.close();
}

export default db;
