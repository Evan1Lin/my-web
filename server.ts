import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { fileURLToPath } from "url";
import multer from "multer";
import * as XLSX from "xlsx";
import dotenv from "dotenv";
import bcrypt from "bcryptjs";
import jwt from "jsonwebtoken";
import { randomUUID } from "crypto";

import {
  getAllIssues,
  insertIssue,
  bulkInsertIssues,
  deleteAllIssues,
  updateIssue,
  deleteIssue,
  getKPIStats,
  getTrendData,
  getModelRanking,
  getCauseDistribution,
  getPerformanceData,
  getUserByUsername,
  updateUserSessionId,
  type QualityIssue,
  type IssueFilters,
} from "./db.ts";

dotenv.config();

const __filename = fileURLToPath(import.meta.url);
const __dirname = path.dirname(__filename);

// Multer for file uploads (in memory)
const upload = multer({ storage: multer.memoryStorage() });

export async function startServer(portOverride?: number): Promise<number> {
  const app = express();
  const PORT = portOverride || parseInt(process.env.PORT || "3000", 10);

  // Middleware
  app.use(express.json({ limit: "50mb" }));
  app.use(express.urlencoded({ extended: true, limit: "50mb" }));

  // ===================== API Routes =====================

  const JWT_SECRET = process.env.JWT_SECRET || "quality_bi_super_secret_key";

  // Authentication Middleware
  app.use("/api", (req, res, next) => {
    // Skip auth for health and login
    if (req.path === "/health" || req.path === "/auth/login") {
      return next();
    }
    const authHeader = req.headers["authorization"];
    const token = authHeader && authHeader.split(" ")[1];
    
    if (!token) {
      return res.status(401).json({ success: false, error: "未授权：请先登录" });
    }
    
    jwt.verify(token, JWT_SECRET, (err, decoded: any) => {
      if (err) {
        return res.status(403).json({ success: false, error: "登录已过期或无效" });
      }
      
      // Concurrent session check
      const user = getUserByUsername(decoded.username);
      if (!user || user.last_session_id !== decoded.sessionId) {
        return res.status(401).json({ 
          success: false, 
          error: "SESSION_EXPIRED_CONCURRENT", 
          message: "该账号已被他人登录！" 
        });
      }

      (req as any).user = decoded;
      next();
    });
  });

  /** POST /api/auth/login */
  app.post("/api/auth/login", (req, res) => {
    try {
      const { username, password } = req.body;
      if (!username || !password) {
        return res.status(400).json({ success: false, error: "必须提供用户名和密码" });
      }
      
      const user = getUserByUsername(username);
      if (!user) {
        return res.status(401).json({ success: false, error: "用户名或密码错误" });
      }
      
      const valid = bcrypt.compareSync(password, user.password_hash);
      if (!valid) {
        return res.status(401).json({ success: false, error: "用户名或密码错误" });
      }
      
      const sessionId = randomUUID();
      updateUserSessionId(user.username, sessionId);
      
      const token = jwt.sign({ id: user.id, username: user.username, sessionId }, JWT_SECRET, { expiresIn: "7d" });
      res.json({ success: true, token, username: user.username });
    } catch (error: any) {
      console.error("Login error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** Health check */
  app.get("/api/health", (_req, res) => {
    res.json({ status: "ok", message: "Quality BI Backend is running", timestamp: new Date().toISOString() });
  });

  /** GET /api/auth/session-status - Simple heartbeat for polling */
  app.get("/api/auth/session-status", (_req, res) => {
    res.json({ success: true });
  });

  /** GET /api/issues — Query issues with optional filters */
  app.get("/api/issues", (req, res) => {
    try {
      const filters: IssueFilters = {
        month: req.query.month ? (req.query.month === "all" ? "all" : parseInt(req.query.month as string, 10)) : "all",
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
        search: (req.query.search as string) || undefined,
      };
      const issues = getAllIssues(filters);
      res.json({ success: true, data: issues, total: issues.length });
    } catch (error: any) {
      console.error("GET /api/issues error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** POST /api/issues — Insert a single issue */
  app.post("/api/issues", (req, res) => {
    try {
      const issue = insertIssue(req.body);
      res.json({ success: true, data: issue });
    } catch (error: any) {
      console.error("POST /api/issues error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** POST /api/issues/bulk — Bulk insert issues */
  app.post("/api/issues/bulk", (req, res) => {
    try {
      const issues: Omit<QualityIssue, "id" | "createdAt">[] = req.body;
      if (!Array.isArray(issues) || issues.length === 0) {
        return res.status(400).json({ success: false, error: "请提供一个非空的数据数组" });
      }
      const count = bulkInsertIssues(issues);
      res.json({ success: true, inserted: count });
    } catch (error: any) {
      console.error("POST /api/issues/bulk error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** PUT /api/issues/:id — Update a single issue */
  app.put("/api/issues/:id", (req, res) => {
    try {
      const id = parseInt(req.params.id, 10);
      const updated = updateIssue(id, req.body);
      if (updated) {
        res.json({ success: true });
      } else {
        res.status(404).json({ success: false, error: "记录未找到" });
      }
    } catch (error: any) {
      console.error("PUT /api/issues error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** DELETE /api/issues — Delete all issues (reset) */
  app.delete("/api/issues", (_req, res) => {
    try {
      const deleted = deleteAllIssues();
      res.json({ success: true, deleted });
    } catch (error: any) {
      console.error("DELETE /api/issues error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** DELETE /api/issues/:id — Delete a single issue */
  app.delete("/api/issues/:id", (req, res) => {
    try {
      const id = parseInt(req.params.id, 10);
      const deleted = deleteIssue(id);
      if (deleted) {
        res.json({ success: true });
      } else {
        res.status(404).json({ success: false, error: "记录未找到" });
      }
    } catch (error: any) {
      console.error("DELETE /api/issues/:id error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** GET /api/kpi — Get KPI aggregated stats */
  app.get("/api/kpi", (req, res) => {
    try {
      const filters: IssueFilters = {
        month: req.query.month ? (req.query.month === "all" ? "all" : parseInt(req.query.month as string, 10)) : "all",
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
      };
      const stats = getKPIStats(filters);
      res.json({ success: true, data: stats });
    } catch (error: any) {
      console.error("GET /api/kpi error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** GET /api/trend — Get monthly trend data */
  app.get("/api/trend", (req, res) => {
    try {
      const filters: IssueFilters = {
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
      };
      const trend = getTrendData(filters);
      res.json({ success: true, data: trend });
    } catch (error: any) {
      console.error("GET /api/trend error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** GET /api/ranking — Get model ranking */
  app.get("/api/ranking", (req, res) => {
    try {
      const filters: IssueFilters = {
        month: req.query.month ? (req.query.month === "all" ? "all" : parseInt(req.query.month as string, 10)) : "all",
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
      };
      const ranking = getModelRanking(filters);
      res.json({ success: true, data: ranking });
    } catch (error: any) {
      console.error("GET /api/ranking error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** GET /api/distribution — Get cause distribution */
  app.get("/api/distribution", (req, res) => {
    try {
      const filters: IssueFilters = {
        month: req.query.month ? (req.query.month === "all" ? "all" : parseInt(req.query.month as string, 10)) : "all",
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
      };
      const dist = getCauseDistribution(filters);
      res.json({ success: true, data: dist });
    } catch (error: any) {
      console.error("GET /api/distribution error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** GET /api/performance — Get performance data */
  app.get("/api/performance", (req, res) => {
    try {
      const filters: IssueFilters = {
        month: req.query.month ? (req.query.month === "all" ? "all" : parseInt(req.query.month as string, 10)) : "all",
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
      };
      const perf = getPerformanceData(filters);
      res.json({ success: true, data: perf });
    } catch (error: any) {
      console.error("GET /api/performance error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** POST /api/import — Server-side Excel import */
  app.post("/api/import", upload.single("file"), (req, res) => {
    try {
      if (!req.file) {
        return res.status(400).json({ success: false, error: "请上传 Excel 文件" });
      }

      const wb = XLSX.read(req.file.buffer, { type: "buffer" });
      const wsname = wb.SheetNames[0];
      const ws = wb.Sheets[wsname];
      const rawData: any[] = XLSX.utils.sheet_to_json(ws);

      const processedData = rawData.map((row) => {
        const normalizeDateText = (value: any) => {
          if (!value) return "";
          if (value instanceof Date) return value.toISOString().slice(0, 10);
          return String(value).trim();
        };

        const parseMonthFromDateLike = (value: any) => {
          if (!value) return 1;
          if (value instanceof Date) return value.getMonth() + 1;
          const text = String(value);
          const m = text.match(/(\d{4})年(\d{1,2})月(\d{1,2})日/);
          if (m) return Math.min(12, Math.max(1, Number(m[2])));
          const d = new Date(text);
          if (!isNaN(d.getTime())) return d.getMonth() + 1;
          return 1;
        };

        const productModelPath = String(row["产品型号"] || row["产品型号名称"] || "").trim();
        const segments = productModelPath
          .split("/")
          .map((s: string) => s.trim())
          .filter(Boolean);
        const level1 = segments[0] || "";

        const productLine = (() => {
          if (level1.includes("机械臂")) return "roboticArm";
          if (level1.includes("机器人")) return "robot";
          if (level1.includes("关节")) return "joint";
          if (level1.includes("其他")) return "others";
          return "others";
        })();

        const issueQuantity = Number(row["问题数量"]) || 0;
        const closedQuantity = Number(row["问题关闭数量"]) || 0;

        const oobRaw = row["是否开箱损"];
        const oobText = typeof oobRaw === "string" ? oobRaw.trim() : oobRaw;
        const oob = (() => {
          if (oobText === true || oobText === "是" || oobText === "Y") return 1;
          if (typeof oobText === "string") {
            if (oobText.includes("非开箱损")) return 0;
            if (oobText.includes("开箱损")) return 1;
          }
          return 0;
        })();

        return {
          month: parseMonthFromDateLike(row["创建时间"]),
          customerName: row["客户名称"] || row["标题"] || row["标题_1"] || "未知客户",
          productQuantity: Number(row["产品数量"]) || 0,
          productModelPath: productModelPath || "未知型号",
          model: segments[segments.length - 1] || productModelPath || "未知型号",
          initialDept: String(row["问题初判主责部门"] || "").trim(),
          mainDept: String(row["问题主责部门"] || "").trim(),
          complaintDate: normalizeDateText(row["创建时间"]),
          snDate: normalizeDateText(row["SN日期"]),
          cause: row["根因分类"] || "未分类",
          analysisType: String(row["问题分析类型"] || "").trim(),
          dept: String(row["问题主责部门"] || row["产品归属"] || "未指定").trim(),
          productLine,
          issueQuantity,
          closedQuantity,
          closed: issueQuantity > 0 && closedQuantity >= issueQuantity ? 1 : 0,
          oob,
          creator: row["创建人"] || "System",
        };
      });

      const count = bulkInsertIssues(processedData);
      res.json({ success: true, inserted: count, message: `成功导入 ${count} 条数据` });
    } catch (error: any) {
      console.error("POST /api/import error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  /** GET /api/export — Export data as Excel file */
  app.get("/api/export", (req, res) => {
    try {
      const filters: IssueFilters = {
        month: req.query.month ? (req.query.month === "all" ? "all" : parseInt(req.query.month as string, 10)) : "all",
        productLine: (req.query.productLine as string) || "all",
        cause: (req.query.cause as string) || "all",
        dept: (req.query.dept as string) || "all",
        oob: (req.query.oob as string) || "all",
      };
      const issues = getAllIssues(filters);

      // Map to Chinese headers for export
      const exportData = issues.map((row) => ({
        月份: row.month,
        客户名称: row.customerName,
        产品型号: row.model,
        根因分类: row.cause,
        产品归属: row.dept,
        问题分析: row.productLine,
        问题数量: row.issueQuantity,
        是否完成: row.closed ? "是" : "否",
        是否开箱损: row.oob ? "是" : "否",
        创建人: row.creator,
        创建时间: row.createdAt,
      }));

      const wb = XLSX.utils.book_new();
      const ws = XLSX.utils.json_to_sheet(exportData);

      // Set column widths
      ws["!cols"] = [
        { wch: 6 },  // 月份
        { wch: 20 }, // 客户名称
        { wch: 18 }, // 产品型号
        { wch: 22 }, // 根因分类
        { wch: 20 }, // 产品归属
        { wch: 15 }, // 问题分析
        { wch: 8 },  // 问题数量
        { wch: 8 },  // 是否完成
        { wch: 10 }, // 是否开箱损
        { wch: 12 }, // 创建人
        { wch: 20 }, // 创建时间
      ];

      XLSX.utils.book_append_sheet(wb, ws, "质量数据");
      const buffer = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });

      const filename = `质量数据报表_${new Date().toISOString().slice(0, 10)}.xlsx`;
      res.setHeader("Content-Disposition", `attachment; filename*=UTF-8''${encodeURIComponent(filename)}`);
      res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
      res.send(buffer);
    } catch (error: any) {
      console.error("GET /api/export error:", error);
      res.status(500).json({ success: false, error: error.message });
    }
  });

  // ===================== Vite / Static =====================

  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    app.use(express.static(path.join(__dirname, "dist")));
    app.get("*", (_req, res) => {
      res.sendFile(path.join(__dirname, "dist", "index.html"));
    });
  }

  return new Promise<number>((resolve, reject) => {
    const server = app.listen(PORT, "0.0.0.0", () => {
      const addr = server.address();
      const actualPort = typeof addr === 'object' && addr ? addr.port : PORT;
      console.log(`✅ Quality BI Server running on http://localhost:${actualPort}`);
      resolve(actualPort);
    });
    server.on('error', (err: any) => {
      if (err.code === 'EADDRINUSE') {
        console.log(`Port ${PORT} in use, trying ${PORT + 1}...`);
        server.close();
        startServer(PORT + 1).then(resolve).catch(reject);
      } else {
        reject(err);
      }
    });
  });
}

// Auto-start when run directly (not imported by Electron)
const isMainModule = !process.env.ELECTRON_RUN_AS_NODE && process.argv[1] && (
  process.argv[1].endsWith('server.ts') || process.argv[1].endsWith('server.js')
);
if (isMainModule) {
  startServer();
}
