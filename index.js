// ============================================================
// OT LINE Bot + LIFF Server — index.js  (v1.1 — B+ patches)
// บริษัท Adrun
// LINE Messaging API (Bot) + LIFF REST API + Static file serve
// ============================================================
// CHANGES vs v1.0:
//   • Employees / Holidays / OT_Records / Edit_Requests อ่านจาก row 3 (ข้าม 2 header rows)
//   • เพิ่มคอลัมน์ userId (D) ใน Employees → match ด้วย userId เป็นหลัก
//   • Auto-bind userId ตอน user แรก login (จับคู่ด้วย displayName ครั้งแรก)
//   • เพิ่ม endpoints: DELETE employees, PUT/DELETE holidays
//   • Auto-set "วันในสัปดาห์" ตอนเพิ่ม/แก้วันหยุด
// ============================================================
require("dotenv").config();
const express    = require("express");
const line       = require("@line/bot-sdk");
const { google } = require("googleapis");
const path       = require("path");
const crypto     = require("crypto");
// ★ v1.28: ใช้ xlsx-js-style (drop-in replacement ของ xlsx ที่รองรับ border + alignment)
const XLSX       = require("xlsx-js-style");

// ★ v1.28: helper — ใส่ border + center alignment ให้ทุก cell ใน worksheet
//   - แถวแรก (header): bg เขียว + text ขาว + bold
//   - cell อื่น ๆ: border บาง + center
function styleSheet(ws, opts = {}) {
  if (!ws || !ws["!ref"]) return ws;
  const range = XLSX.utils.decode_range(ws["!ref"]);
  const headerRowIdx = opts.headerRow !== undefined ? opts.headerRow : 0; // default = แถว 0
  const titleRowIdx  = opts.titleRow;  // optional — แถวชื่อรายงาน (merge cells)
  const thinBorder = {
    top:    { style: "thin", color: { rgb: "888888" } },
    bottom: { style: "thin", color: { rgb: "888888" } },
    left:   { style: "thin", color: { rgb: "888888" } },
    right:  { style: "thin", color: { rgb: "888888" } },
  };
  for (let R = range.s.r; R <= range.e.r; R++) {
    for (let C = range.s.c; C <= range.e.c; C++) {
      const cellRef = XLSX.utils.encode_cell({ r: R, c: C });
      const cell = ws[cellRef];
      if (!cell) continue; // ข้าม cell ว่าง (ไม่ใส่ border)
      const isHeader = R === headerRowIdx;
      const isTitle  = R === titleRowIdx;
      const isNumber = typeof cell.v === "number";
      cell.s = {
        font: {
          name: "TH SarabunPSK",
          sz: isTitle ? 16 : (isHeader ? 12 : 11),
          bold: isHeader || isTitle,
          color: { rgb: isHeader ? "FFFFFF" : "222222" },
        },
        alignment: {
          horizontal: isTitle ? "center" : (isNumber ? "center" : "center"),
          vertical: "center",
          wrapText: true,
        },
        border: thinBorder,
        fill: isHeader
          ? { patternType: "solid", fgColor: { rgb: "2E7D32" } }     // เขียวเข้ม
          : isTitle
            ? { patternType: "solid", fgColor: { rgb: "E8F5E9" } }   // เขียวอ่อน
            : { patternType: "solid", fgColor: { rgb: "FFFFFF" } },
      };
      // number format
      if (isNumber && !isHeader && !isTitle) {
        cell.z = "#,##0";
      }
    }
  }
  // row heights
  ws["!rows"] = ws["!rows"] || [];
  for (let R = range.s.r; R <= range.e.r; R++) {
    ws["!rows"][R] = { hpt: R === titleRowIdx ? 26 : 22 };
  }
  return ws;
}

// ★ v1.29: page setup ให้พิมพ์พอดี A4
//   - landscape (เพราะตารางกว้าง)
//   - fit-to-width = 1 หน้า, height = auto
//   - margin บาง 0.4 นิ้ว
//   - จัดกลางหน้า (horizontalCentered)
function applyA4Print(ws, opts = {}) {
  if (!ws) return ws;
  const orientation = opts.orientation || "landscape";
  ws["!pageSetup"] = {
    paperSize: 9,           // 9 = A4
    orientation,            // 'landscape' | 'portrait'
    fitToWidth: 1,
    fitToHeight: 0,         // 0 = auto (กี่หน้าก็ได้)
    scale: 100,
    horizontalDpi: 300,
    verticalDpi: 300,
  };
  ws["!margins"] = {
    left: 0.4, right: 0.4,
    top: 0.5, bottom: 0.5,
    header: 0.3, footer: 0.3,
  };
  // จัดกลางหน้า (horizontal) + fit page
  ws["!printOptions"] = {
    horizontalCentered: true,
  };
  // sheetPr.pageSetUpPr.fitToPage = true
  ws["!sheetPr"] = {
    pageSetUpPr: { fitToPage: true },
  };
  return ws;
}

// ★ v1.29: ตั้ง print area + repeat header rows ให้ workbook
function setPrintTitles(wb, sheetName, headerRowsZeroBased) {
  if (!headerRowsZeroBased || headerRowsZeroBased < 0) return;
  wb.Workbook = wb.Workbook || {};
  wb.Workbook.Names = wb.Workbook.Names || [];
  const sheetIdx = wb.SheetNames.indexOf(sheetName);
  if (sheetIdx < 0) return;
  // ใช้ชื่อ sheet ห่อด้วย ' (สำหรับชื่อภาษาไทย)
  wb.Workbook.Names.push({
    Name: "_xlnm.Print_Titles",
    Sheet: sheetIdx,
    Ref: `'${sheetName}'!$1:$${headerRowsZeroBased + 1}`,
  });
}

// ── LINE Config ──────────────────────────────────────────────
const lineConfig = {
  channelAccessToken: process.env.LINE_TOKEN,
  channelSecret:      process.env.LINE_SECRET,
};
const client = new line.Client(lineConfig);
const app    = express();

// ── Webhook ต้องมาก่อน express.json() เสมอ ──────────────────
// ★ v1.30 SEC-01: verify HMAC signature ก่อนตอบ 200 — กันคนยิงปลอม
app.post("/webhook", express.raw({ type: "*/*" }), async (req, res) => {
  const sig  = req.headers["x-line-signature"];
  const body = req.body;
  if (!sig || !body) return res.status(400).end();

  let hash;
  try {
    hash = crypto.createHmac("SHA256", process.env.LINE_SECRET)
                 .update(body).digest("base64");
  } catch (e) {
    console.error("HMAC compute failed:", e.message);
    return res.status(500).end();
  }

  if (sig !== hash) {
    console.error("❌ Webhook signature mismatch — drop");
    return res.status(401).end();
  }

  // ✅ verified — respond OK then process events async (LINE retries ถ้า > 1 sec)
  res.status(200).end();

  let events = [];
  try {
    events = JSON.parse(body.toString()).events || [];
  } catch (e) {
    console.error("Webhook JSON parse failed:", e.message);
    return;
  }
  console.log(`📋 Events: ${events.length}`, events.map(e => `${e.type}:${e.message?.text || ""}`));
  await Promise.all(
    events.map(e => handleBotEvent(e).catch(err => console.error("handleBotEvent:", err.message)))
  );
});

// ── Middleware ───────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, "public")));   // serve LIFF HTML

// ── Google Sheets ────────────────────────────────────────────
const SHEET_ID = process.env.GOOGLE_SHEET_ID;

async function getSheetsClient() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON),
    scopes: ["https://www.googleapis.com/auth/spreadsheets"],
  });
  return google.sheets({ version: "v4", auth });
}

// ══════════════════════════════════════════════════════════════
// ★ v1.30 SEC-02: AUTH HELPERS
//   - getAdminIds()         : list ของ admin LINE IDs
//   - requireAdmin          : middleware เช็ค header x-line-user-id
//   - verifyLineIdToken     : verify ID token กับ LINE oauth2 endpoint
// ══════════════════════════════════════════════════════════════
function getAdminIds() {
  return (process.env.ADMIN_LINE_IDS || "").split(",").map(s => s.trim()).filter(Boolean);
}

// ★ v1.32: helper — เช็คสิทธิ์ admin/supervisor จากทั้ง env + sheet
//   admin   = (uid in ADMIN_LINE_IDS env) OR (employee role="admin")
//   supervisor = (employee role="supervisor") OR is admin (admin includes supervisor power)
async function getRoleForUser(uid) {
  if (!uid) return "employee";
  if (getAdminIds().includes(uid)) return "admin";
  try {
    const sheets = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const emp = employees.find(e => e.userId === uid);
    if (!emp) return "employee";
    if (emp.role === "admin") return "admin";
    if (emp.role === "supervisor") return "supervisor";
    return "employee";
  } catch (_) { return "employee"; }
}

// ★ v1.32 fix: เช็ค admin จาก env หรือ sheet role="admin" (เดิม env เท่านั้น)
async function requireAdmin(req, res, next) {
  try {
    const uid = (
      req.headers["x-line-user-id"] ||
      req.query._uid ||
      ""
    ).toString().trim();
    if (!uid) return res.status(403).json({ error: "Forbidden — admin only" });
    if (getAdminIds().includes(uid)) return next();
    // เช็ค sheet role
    const sheets = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const emp = employees.find(e => e.userId === uid);
    if (emp && emp.role === "admin") return next();
    return res.status(403).json({ error: "Forbidden — admin only" });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}

// ★ v1.32: middleware สำหรับ supervisor — ผ่านได้ถ้า env-admin หรือ sheet role ∈ {admin, supervisor}
async function requireSupervisor(req, res, next) {
  try {
    const uid = (req.headers["x-line-user-id"] || req.query._uid || "").toString().trim();
    if (!uid) return res.status(403).json({ error: "Forbidden — supervisor only" });
    if (getAdminIds().includes(uid)) return next();  // admin = supervisor power
    const sheets = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const emp = employees.find(e => e.userId === uid);
    if (emp && (emp.role === "supervisor" || emp.role === "admin")) return next();
    return res.status(403).json({ error: "Forbidden — supervisor only" });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
}

// verify LIFF ID token กับ LINE oauth2 endpoint
// (ใช้ env LIFF_CHANNEL_ID — เลขช่อง LINE Login channel ที่ผูกกับ LIFF)
async function verifyLineIdToken(idToken) {
  if (!idToken) return null;
  const channelId = process.env.LIFF_CHANNEL_ID;
  if (!channelId) {
    console.warn("⚠️ LIFF_CHANNEL_ID ยังไม่ได้ตั้ง → ข้าม verify (ไม่ปลอดภัย)");
    return null;
  }
  try {
    const params = new URLSearchParams({ id_token: idToken, client_id: channelId });
    const r = await fetch("https://api.line.me/oauth2/v2.1/verify", {
      method: "POST",
      headers: { "Content-Type": "application/x-www-form-urlencoded" },
      body: params.toString(),
    });
    if (!r.ok) {
      const t = await r.text().catch(() => "");
      console.warn("verifyLineIdToken failed:", r.status, t.slice(0, 200));
      return null;
    }
    const data = await r.json();
    return { sub: data.sub, name: data.name, email: data.email || "" };
  } catch (e) {
    console.error("verifyLineIdToken error:", e.message);
    return null;
  }
}

// ══════════════════════════════════════════════════════════════
// ★ v1.31 D3: simple in-memory rate limiter (per userId / fallback IP)
//   - 30 requests / 60 sec → 429 Too Many Requests
//   - cleanup ทุก 5 นาที กัน Map โต
// ══════════════════════════════════════════════════════════════
const __rl = new Map();           // key → array of timestamps (ms)
const RL_WINDOW_MS = 60_000;
const RL_MAX = 30;
function rateLimitByUser(req, res, next) {
  const key = (
    req.headers["x-line-user-id"] ||
    req.query._uid ||
    req.ip ||
    "anon"
  ).toString();
  const now = Date.now();
  const arr = (__rl.get(key) || []).filter(t => now - t < RL_WINDOW_MS);
  if (arr.length >= RL_MAX) {
    return res.status(429).json({ error: "ส่งคำขอเร็วเกินไป รอสักครู่แล้วลองใหม่" });
  }
  arr.push(now);
  __rl.set(key, arr);
  next();
}
setInterval(() => {
  const now = Date.now();
  for (const [k, arr] of __rl) {
    if (arr.every(t => now - t > RL_WINDOW_MS * 5)) __rl.delete(k);
  }
}, 5 * 60_000).unref();

// ══════════════════════════════════════════════════════════════
// ★ v1.30 BUG-04: in-memory mutex per (name|date)
//   กัน race ตอน 2 client ยิง POST /api/ot พร้อมกัน
//   (Railway มักรัน 1 instance — ถ้า scale-out ต้องเปลี่ยนเป็น distributed lock)
// ══════════════════════════════════════════════════════════════
const __otMutexes = new Map();
function withOTMutex(key, fn) {
  const prev = __otMutexes.get(key) || Promise.resolve();
  let release;
  const next = new Promise(r => { release = r; });
  __otMutexes.set(key, prev.then(() => next).catch(() => next));
  return prev
    .catch(() => {})
    .then(async () => {
      try { return await fn(); }
      finally {
        release();
        if (__otMutexes.get(key) === next) __otMutexes.delete(key);
      }
    });
}

// ── OT Rules ─────────────────────────────────────────────────
const MAX_OT_PER_DAY     = 5;
const WEEKDAY_MULTIPLIER = 1;  // ★ v1.34: hourlyRate ที่กรอกใน Employees คืออัตราค่า OT ต่อ ชม. อยู่แล้ว ไม่ต้องคูณ 1.5 ซ้ำ

// ★ v1.3: เวลางานปกติ จันทร์–เสาร์ (ห้ามลง OT ทับ)
const WORK_START_MIN     = 8 * 60 + 30;   // 08:30 = 510
const WORK_END_MIN       = 17 * 60 + 30;  // 17:30 = 1050
// ★ v1.27: เสาร์ครึ่งวันสำหรับพนักงานบางคน — ทำถึง 12:00 → OT เริ่มได้ตั้งแต่ 12:00
const WORK_END_MIN_SAT_HALF = 12 * 60;    // 12:00 = 720

// ★ v1.5: OT day = 06:00 ถึง 06:00 ของวันถัดไป (24 ชม.)
const OT_DAY_START_MIN = 6 * 60;                  // 06:00 = 360
const OT_DAY_END_MIN   = OT_DAY_START_MIN + 1440; // 06:00 next day = 1800

// ★ data ทุก sheet เริ่มที่ row 3 (มี 2 header rows)
const DATA_START_ROW = 3;

// Helper: idx (0-based ใน array) → row จริง (1-indexed) ใน sheet
const idxToRow = (idx) => idx + DATA_START_ROW;

// ══════════════════════════════════════════════════════════════
// LIFF REST API ENDPOINTS
// ══════════════════════════════════════════════════════════════

// ── GET /api/config — ส่ง LIFF ID ไปให้ front-end ──────────
app.get("/api/config", (_, res) => {
  res.json({ liffId: process.env.LIFF_ID });
});

// ── GET /api/me — ★ v1.30 SEC-03: ใช้ Authorization Bearer <idToken>
// ฝั่ง client ส่ง: Authorization: Bearer <liff.getIDToken()>
// ถ้า verify ผ่าน → ใช้ userId+displayName จาก token (ปลอม-ไม่-ได้)
// ถ้าไม่ส่ง token → fallback ไปอ่าน query string (legacy, log warning)
app.get("/api/me", async (req, res) => {
  let userId, displayName;
  let verified = false;

  // ★ v1.32 fix: ลำดับการหา userId
  //   1) ID token (verify ก่อน) → verified=true, auto-bind ทำได้
  //   2) ถ้า token หมดอายุ/fail → fall back x-line-user-id header (verified=false, ไม่ auto-bind)
  //   3) สุดท้าย legacy query string
  const auth = req.headers.authorization || "";
  const idToken = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";
  if (idToken) {
    const claims = await verifyLineIdToken(idToken);
    if (claims) {
      userId = claims.sub;
      displayName = claims.name || "";
      verified = true;
    } else {
      console.warn("⚠️ /api/me ID token invalid — fall back to header");
    }
  }
  if (!userId) {
    userId = (req.headers["x-line-user-id"] || req.query.userId || "").toString().trim();
    displayName = req.query.displayName || "";
  }

  try {
    const sheets    = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const admins    = getAdminIds();

    // 1) จับคู่ด้วย userId ก่อน (น่าเชื่อถือสุด)
    let emp = employees.find(e => e.userId && e.userId === userId);
    let matchedBy = emp ? "userId" : null;

    // 2) ถ้าไม่เจอ → fallback หาด้วย displayName
    //    ★ SEC: auto-bind จะทำเฉพาะกรณี verified=true (ID token) เพื่อกันปลอมตัว
    if (!emp && displayName) {
      emp = employees.find(e => e.name === displayName);
      if (emp) {
        matchedBy = "displayName";
        if (!emp.userId && userId && verified) {
          const row = idxToRow(emp.idx);
          try {
            await sheets.spreadsheets.values.update({
              spreadsheetId: SHEET_ID,
              range: `Employees!D${row}`,
              valueInputOption: "USER_ENTERED",
              resource: { values: [[userId]] },
            });
            console.log(`🔗 Auto-bound userId ${userId} → ${emp.name}`);
            matchedBy = "displayName+autoBind";
          } catch (err) {
            console.error("auto-bind failed:", err.message);
          }
        } else if (!emp.userId && userId && !verified) {
          // legacy mode — DO NOT auto-bind (ป้องกัน impersonation)
          matchedBy = "displayName(unverified)";
        }
      }
    }

    const isAdmin = admins.includes(userId) || emp?.role === "admin";
    // ★ v1.32: คำนวณ role — admin > supervisor > employee
    let role = "employee";
    if (isAdmin) role = "admin";
    else if (emp?.role === "supervisor") role = "supervisor";

    res.json({
      found:       !!emp,
      name:        emp?.name || displayName || "",
      userId,
      isAdmin,
      role,                       // ★ v1.32: "admin" | "supervisor" | "employee"
      hourlyRate:  emp?.hourlyRate  || 0,
      holidayFlat: emp?.holidayFlat || 0,
      matchedBy,
      verified,  // ★ ให้ client เช็คได้ว่า ID token verify ผ่านไหม
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/records?name=xxx&month=04&year=2568 ──────────────
app.get("/api/records", async (req, res) => {
  const { name, month, year } = req.query;
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    const rows   = all.filter(r =>
      r.name === name &&
      (!month || r.date.includes(`/${month}/${year}`))
    );
    res.json({ records: rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/employees ──────────────────────────────────────
app.get("/api/employees", async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    res.json({ employees: await getEmployees(sheets) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/employees — เพิ่มพนักงาน (Admin) ──────────────
app.post("/api/employees", requireAdmin, async (req, res) => {
  const { name, hourlyRate, holidayFlat, userId, outProvinceFlat, travelAllowance, socialSecurity, satHalfDay, role, salary } = req.body;
  if (!name || !String(name).trim()) return res.status(400).json({ error: "กรุณากรอกชื่อพนักงาน" });
  try {
    const sheets = await getSheetsClient();
    // ★ v1.30: ป้องกันชื่อซ้ำ
    const existing = await getEmployees(sheets);
    if (existing.some(e => e.name === String(name).trim())) {
      return res.status(400).json({ error: `มีพนักงานชื่อ "${name}" อยู่แล้วในระบบ` });
    }
    const safeRole = ["admin","supervisor"].includes((role||"").toLowerCase()) ? role.toLowerCase() : "";
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "Employees!A:J",  // ★ v1.33: + col J salary
      valueInputOption: "USER_ENTERED",
      resource: { values: [[
        name, hourlyRate, holidayFlat, userId || "",
        Number(outProvinceFlat) || 0,
        Number(travelAllowance) || 0,
        Number(socialSecurity) || 0,
        satHalfDay ? "TRUE" : "",
        safeRole,
        Number(salary) || 0,        // ★ v1.33
      ]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── PUT /api/employees/:idx — แก้ไขพนักงาน (Admin) ──────────
app.put("/api/employees/:idx", requireAdmin, async (req, res) => {
  const row = idxToRow(Number(req.params.idx));
  const { name, hourlyRate, holidayFlat, userId, outProvinceFlat, travelAllowance, socialSecurity, satHalfDay, role, salary } = req.body;
  try {
    const sheets = await getSheetsClient();
    const safeRole = ["admin","supervisor"].includes((role||"").toLowerCase()) ? role.toLowerCase() : "";
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Employees!A${row}:J${row}`,  // ★ v1.33: + col J salary
      valueInputOption: "USER_ENTERED",
      resource: { values: [[
        name, hourlyRate, holidayFlat, userId || "",
        Number(outProvinceFlat) || 0,
        Number(travelAllowance) || 0,
        Number(socialSecurity) || 0,
        satHalfDay ? "TRUE" : "",
        safeRole,
        Number(salary) || 0,        // ★ v1.33
      ]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── DELETE /api/employees/:idx — ลบพนักงาน (Admin) ★ NEW ────
app.delete("/api/employees/:idx", requireAdmin, async (req, res) => {
  const row = idxToRow(Number(req.params.idx));
  try {
    const sheets = await getSheetsClient();
    await deleteRow(sheets, "Employees", row);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/holidays ───────────────────────────────────────
app.get("/api/holidays", async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    res.json({ holidays: await getHolidayList(sheets) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/holidays — เพิ่มวันหยุด (Admin) ───────────────
app.post("/api/holidays", requireAdmin, async (req, res) => {
  const { date, name: hName } = req.body;
  if (!date || !/^\d{2}\/\d{2}\/\d{4}$/.test(date)) {
    return res.status(400).json({ error: "รูปแบบวันที่ผิด ต้องเป็น dd/mm/yyyy (พ.ศ.)" });
  }
  try {
    const sheets = await getSheetsClient();
    // ★ v1.30: ป้องกันวันหยุดซ้ำ
    const existing = await getHolidayList(sheets);
    if (existing.some(h => h.date === date)) {
      return res.status(400).json({ error: `วันที่ ${date} มีในรายการวันหยุดแล้ว` });
    }
    const dow    = getDowFullThai(date);
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "Holidays!A:C",
      valueInputOption: "USER_ENTERED",
      resource: { values: [[date, dow, hName]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── PUT /api/holidays/:idx — แก้ไขวันหยุด (Admin) ★ NEW ─────
app.put("/api/holidays/:idx", requireAdmin, async (req, res) => {
  const row = idxToRow(Number(req.params.idx));
  const { date, name: hName } = req.body;
  try {
    const sheets = await getSheetsClient();
    const dow    = getDowFullThai(date);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Holidays!A${row}:C${row}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[date, dow, hName]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── DELETE /api/holidays/:idx — ลบวันหยุด (Admin) ★ NEW ─────
app.delete("/api/holidays/:idx", requireAdmin, async (req, res) => {
  const row = idxToRow(Number(req.params.idx));
  try {
    const sheets = await getSheetsClient();
    await deleteRow(sheets, "Holidays", row);
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/ot — บันทึก OT ────────────────────────────────
// ★ v1.30 BUG-04: wrap ด้วย mutex ต่อ (name|date) กัน race
// ★ v1.31 D3: rate limit
app.post("/api/ot", rateLimitByUser, async (req, res) => {
  const { name, date, startTime, endTime, task, location, otType } = req.body;
  if (!name || !date) return res.status(400).json({ error: "ต้องระบุชื่อและวันที่" });
  try {
    return await withOTMutex(`${name}|${date}`, async () => {
    const sheets    = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const emp       = employees.find(e => e.name === name);
    if (!emp) return res.status(400).json({ error: `ไม่พบชื่อ "${name}" ในระบบ` });

    const holidays = await getHolidayList(sheets);
    const [dd, mm, yy] = date.split("/");
    const jsDate    = new Date(Number(yy) - 543, Number(mm) - 1, Number(dd));
    const dow       = jsDate.getDay();
    const isHolDate = holidays.some(h => h.date === date);
    const isSun     = dow === 0;
    const isHoliday = otType === "holiday" || isHolDate || isSun;

    // ★ v1.25: ทำงานต่างจังหวัด (ค้างคืน) — ไม่ต้องใส่เวลา
    if (otType === "outProvince") {
      // ถ้าวันนั้นเป็นวันหยุด/อาทิตย์ → ใช้ holidayFlat (ตามที่ user ระบุ)
      const useFlat = isHoliday ? emp.holidayFlat : emp.outProvinceFlat;
      const typeLabel = isHoliday
        ? (isSun ? "ต่างจังหวัด+อาทิตย์" : "ต่างจังหวัด+วันหยุด")
        : "ต่างจังหวัด";
      await saveRecord(sheets, {
        name, date, startTime: "-", endTime: "-", hours: 0,
        task, location, otType: typeLabel, pay: useFlat,
      });
      return res.json({ ok: true, hours: 0, pay: useFlat, otType: typeLabel });
    }

    if (isHoliday) {
      const typeLabel = isSun ? "วันอาทิตย์" : isHolDate ? "วันหยุดนักขัตฤกษ์" : "วันหยุด";
      await saveRecord(sheets, {
        name, date, startTime: "-", endTime: "-", hours: 0,
        task, location, otType: typeLabel, pay: emp.holidayFlat,
      });
      return res.json({ ok: true, hours: 0, pay: emp.holidayFlat, otType: typeLabel });
    }

    const hours = calcHours(startTime, endTime);
    if (hours <= 0) return res.status(400).json({ error: "เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น" });

    // ★ v1.6: ลงเวลาได้ทุกช่วง ห้ามแค่ทับเวลางานปกติ
    // ★ v1.27: เสาร์ครึ่งวัน → window 08:30-12:00
    const workEndMin = getWorkEndForEmp(emp, dow);
    if (overlapsWorkHours(startTime, endTime, workEndMin)) {
      return res.status(400).json({ error: `ช่วง ${workWindowLabel(workEndMin)} เป็นเวลางานปกติ ไม่สามารถบันทึก OT ได้` });
    }

    // ★ v1.21: ห้ามบันทึกทับกับ record ในวันเดียวกัน
    const overlap = await findOverlappingRecord(sheets, name, date, startTime, endTime);
    if (overlap) {
      return res.status(400).json({
        error: `ช่วง ${startTime}-${endTime} ทับกับรายการเดิม ${overlap.startTime}-${overlap.endTime} (${overlap.hours} ชม.) ในวันเดียวกัน — ใช้ปุ่ม 🔧 ขอแก้ไขแทน`,
      });
    }

    // ★ v1.2: บันทึกเวลาตามจริง แต่คำนวณค่า OT สูงสุด MAX_OT_PER_DAY ชม./วัน
    const alreadyDay        = await getDayHours(sheets, name, date);
    const remainingPayable  = Math.max(0, MAX_OT_PER_DAY - alreadyDay);
    const payableHours      = Math.min(hours, remainingPayable);
    const pay = Math.round(payableHours * emp.hourlyRate * WEEKDAY_MULTIPLIER);

    await saveRecord(sheets, { name, date, startTime, endTime, hours, task, location, otType: "วันธรรมดา", pay });
    return res.json({ ok: true, hours, payableHours, pay, otType: "วันธรรมดา", capped: payableHours < hours });
    });  // end withOTMutex

  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// ★ v1.35 LEAVE SYSTEM (Phase 1 — backend only, แยกจาก OT_Records)
//   • Sheet: Leave_Records (auto-create บน first use)
//   • Cols: A=name, B=date, C=leaveType, D=reason, E=createdAt
//   • ไม่กระทบ payroll/OT logic เลย — แยกระบบสมบูรณ์
//   ✂️ Rollback: ลบ block นี้ทั้งหมด + ลบ /api/leave* endpoints ด้านล่าง
// ══════════════════════════════════════════════════════════════
const VALID_LEAVE_TYPES = new Set(["ลาป่วย", "ลากิจ", "ลาพักร้อน"]);

// auto-create sheet ถ้ายังไม่มี (idempotent — ทำได้หลายรอบไม่เจ็บ)
async function ensureLeaveSheet(sheets) {
  const meta = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const exists = meta.data.sheets.some(s => s.properties.title === "Leave_Records");
  if (exists) return;
  // เพิ่ม sheet ใหม่ + เขียน header rows
  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: { requests: [{ addSheet: { properties: { title: "Leave_Records" } } }] },
  });
  await sheets.spreadsheets.values.update({
    spreadsheetId: SHEET_ID,
    range: "Leave_Records!A1:E2",
    valueInputOption: "USER_ENTERED",
    resource: { values: [
      ["📋 ระบบลา (Adrun)", "", "", "", ""],
      ["name", "date", "leaveType", "reason", "createdAt"],
    ] },
  });
  console.log("✅ สร้าง sheet Leave_Records แล้ว");
}

async function getAllLeaves(sheets) {
  await ensureLeaveSheet(sheets);
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Leave_Records!A3:E20000",
  });
  return (r.data.values || []).map((row, i) => ({
    idx: i,
    name:      row[0] || "",
    date:      row[1] || "",
    leaveType: row[2] || "",
    reason:    row[3] || "",
    createdAt: row[4] || "",
  }));
}

async function saveLeave(sheets, data) {
  await ensureLeaveSheet(sheets);
  const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "Leave_Records!A:E",
    valueInputOption: "USER_ENTERED",
    resource: { values: [[ data.name, data.date, data.leaveType, data.reason || "", now ]] },
  });
}

// ── POST /api/leave — บันทึกการลา (พนักงานคนใดก็ได้บันทึกของตัวเอง) ──
app.post("/api/leave", rateLimitByUser, async (req, res) => {
  const { name, date, leaveType, reason } = req.body || {};
  if (!name || !date || !leaveType) {
    return res.status(400).json({ error: "กรุณาระบุ name, date, leaveType" });
  }
  if (!/^\d{2}\/\d{2}\/\d{4}$/.test(date)) {
    return res.status(400).json({ error: "date ต้องเป็น DD/MM/YYYY (พ.ศ.)" });
  }
  if (!VALID_LEAVE_TYPES.has(leaveType)) {
    return res.status(400).json({ error: `leaveType ต้องเป็น: ${[...VALID_LEAVE_TYPES].join(" / ")}` });
  }
  // เช็คลาย้อนหลัง
  const [dd, mm, yyyy] = date.split("/").map(Number);
  const leaveDate = new Date(yyyy - 543, mm - 1, dd); leaveDate.setHours(0,0,0,0);
  const today = new Date(); today.setHours(0,0,0,0);
  if (leaveDate < today) {
    return res.status(400).json({ error: "ลาย้อนหลังไม่ได้ — เลือกวันที่ตั้งแต่วันนี้ขึ้นไป" });
  }
  try {
    const sheets = await getSheetsClient();
    const all = await getAllLeaves(sheets);
    const dup = all.find(l => l.name === name && l.date === date);
    if (dup) {
      return res.status(400).json({ error: `วันที่ ${date} คุณบันทึกลาไว้แล้ว (${dup.leaveType})` });
    }
    await saveLeave(sheets, { name, date, leaveType, reason: reason || "" });
    console.log(`📋 Leave: ${name} ${date} ${leaveType} — ${reason || "-"}`);
    return res.json({ ok: true, name, date, leaveType, reason: reason || "" });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ── GET /api/leave?name=&month=&year= — ดึง leave records ──
//   - ไม่มี name → ทุกคน (สำหรับ admin/supervisor)
//   - มี name → เฉพาะคนนั้น
//   - มี month/year → กรองเดือน
app.get("/api/leave", async (req, res) => {
  const { name, month, year } = req.query;
  try {
    const sheets = await getSheetsClient();
    let all = await getAllLeaves(sheets);
    if (name)               all = all.filter(l => l.name === name);
    if (month && year) {
      const monthKey = `/${String(month).padStart(2,"0")}/${year}`;
      all = all.filter(l => l.date && l.date.endsWith(monthKey));
    }
    return res.json({ leaves: all });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ── DELETE /api/leave/:idx — admin ลบ leave record ──
app.delete("/api/leave/:idx", requireAdmin, async (req, res) => {
  const idx = Number(req.params.idx);
  if (!Number.isInteger(idx) || idx < 0) {
    return res.status(400).json({ error: "idx ไม่ถูกต้อง" });
  }
  try {
    const sheets = await getSheetsClient();
    const all = await getAllLeaves(sheets);
    const target = all.find(l => l.idx === idx);
    if (!target) return res.status(404).json({ error: `ไม่พบ leave idx=${idx}` });
    // row จริง = idx + 3 (header rows 1-2 + 1-based)
    await deleteRow(sheets, "Leave_Records", idx + 3);
    console.log(`🗑️  ลบ leave: ${target.name} ${target.date} ${target.leaveType}`);
    return res.json({ ok: true, idx, name: target.name, date: target.date, leaveType: target.leaveType });
  } catch (e) {
    return res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// ★ END LEAVE SYSTEM (Phase 1)
// ══════════════════════════════════════════════════════════════

// ── POST /api/bind-employee — Self-claim ผูกบัญชี LINE ★ v1.13
// ★ v1.32 fix: รับ userId จาก ID token (verified) หรือ x-line-user-id header
//   (ID token ใน LIFF inApp อาจ verify fail — fallback header)
//   ความปลอดภัยยังคง: เช็ค target.userId ว่าง + userId นี้ยังไม่ถูกผูก (logic ด้านล่าง)
app.post("/api/bind-employee", rateLimitByUser, async (req, res) => {
  const { employeeName } = req.body;
  let userId = "";
  const auth = req.headers.authorization || "";
  const idToken = auth.startsWith("Bearer ") ? auth.slice(7).trim() : "";
  if (idToken) {
    const claims = await verifyLineIdToken(idToken);
    if (claims) userId = claims.sub;
  }
  if (!userId) {
    userId = (req.headers["x-line-user-id"] || "").toString().trim();
  }
  if (!userId) return res.status(401).json({ error: "ไม่พบ LINE userId — ลอง logout/login LINE แล้วเปิดใหม่" });
  if (!employeeName) return res.status(400).json({ error: "ข้อมูลไม่ครบ" });
  try {
    const sheets    = await getSheetsClient();
    const employees = await getEmployees(sheets);

    // 1) เช็คว่า userId นี้ผูกกับคนอื่นแล้วหรือยัง
    const existingBind = employees.find(e => e.userId && e.userId === userId);
    if (existingBind) {
      return res.status(400).json({ error: `LINE ID นี้ถูกผูกกับ "${existingBind.name}" แล้ว` });
    }

    // 2) หาเป้าหมาย
    const target = employees.find(e => e.name === employeeName);
    if (!target) return res.status(400).json({ error: `ไม่พบพนักงาน "${employeeName}"` });

    // 3) เช็คว่าเป้าหมายยังไม่ผูกกับใคร
    if (target.userId) {
      return res.status(400).json({ error: `"${employeeName}" ถูกผูกกับ LINE คนอื่นแล้ว` });
    }

    // 4) Bind
    const row = idxToRow(target.idx);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Employees!D${row}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[userId]] },
    });

    console.log(`🔗 Self-claim: ${userId} → ${target.name}`);
    res.json({ ok: true, name: target.name });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/edit-request — ขอแก้ไข OT ────────────────────
// ★ v1.31 D3: rate limit
app.post("/api/edit-request", rateLimitByUser, async (req, res) => {
  const { name, date, recordDesc, note } = req.body;
  try {
    const sheets = await getSheetsClient();

    // ★ v1.15: เช็คว่า record ที่ขอแก้นั้น "จ่ายแล้ว" หรือยัง
    const all = await getAllRecords(sheets);
    const matched = all.find(r => r.name === name && r.date === date);
    if (matched && matched.paidAt) {
      return res.status(400).json({ error: `รายการนี้ทำจ่ายไปแล้ว (รอบ ${matched.paidAt}) ไม่สามารถแก้ไขได้` });
    }

    const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "Edit_Requests!A:F",
      valueInputOption: "USER_ENTERED",
      resource: { values: [[name, date, recordDesc, note, "รอดำเนินการ", now]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// PAYROLL ENDPOINTS (v1.15)
// ══════════════════════════════════════════════════════════════

// ── Helper: เปรียบเทียบวันที่ Thai dd/mm/yyyy ที่ <= cutoff ───
// ★ v1.30 BUG-05: หา set ของพนักงานที่ "รับ travel + SSO ไปแล้ว" ในเดือนของ cutoff
//   - ใช้กับ preview + commit เพื่อกัน duplicate (รายเดือน — จ่ายซ้ำหลายรอบในเดือนเดียวไม่ได้)
async function getEmployeesAlreadyPaidExtrasThisMonth(sheets, allRecords, cutoff) {
  const parts = String(cutoff || "").split("/");
  if (parts.length !== 3) return new Set();
  const monthKey = `/${parts[1]}/${parts[2]}`;
  // อ่าน Payroll_Log
  let logRows = [];
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Payroll_Log!A3:J5000",
    });
    logRows = r.data.values || [];
  } catch (_) { /* tab อาจยังไม่มี */ }
  const activePayIds = new Set();
  // ★ v1.32: dedup includes pending_approval + approved (records lock แล้ว)
  const dedupStatuses = new Set(["pending_approval", "approved", "active"]);
  logRows.forEach(row => {
    const payId    = (row[0] || "").trim();
    const cutoffStr= (row[1] || "").trim();
    const status   = (row[7] || "active").trim();
    if (payId && dedupStatuses.has(status) && cutoffStr.endsWith(monthKey)) {
      activePayIds.add(payId);
    }
  });
  const empsAlreadyPaid = new Set();
  allRecords.forEach(r => {
    if (r.paidAt && activePayIds.has(r.paidAt)) empsAlreadyPaid.add(r.name);
  });
  return empsAlreadyPaid;
}

function isOnOrBeforeCutoff(recordDate, cutoffDate) {
  const [d1,m1,y1] = recordDate.split("/").map(Number);
  const [d2,m2,y2] = cutoffDate.split("/").map(Number);
  const t1 = new Date(y1-543, m1-1, d1).getTime();
  const t2 = new Date(y2-543, m2-1, d2).getTime();
  return t1 <= t2;
}

// ── GET /api/payroll/preview?cutoff=DD/MM/YYYY ────────────
// ดูตัวอย่างก่อนจ่าย — รวม pending records ทั้งหมด ที่ date ≤ cutoff
app.get("/api/payroll/preview", requireAdmin, async (req, res) => {
  const { cutoff } = req.query;
  if (!cutoff || !/^\d{2}\/\d{2}\/\d{4}$/.test(cutoff)) {
    return res.status(400).json({ error: "cutoff ต้องอยู่ในรูป DD/MM/YYYY (พ.ศ.)" });
  }
  try {
    const sheets  = await getSheetsClient();
    const all     = await getAllRecords(sheets);
    // ★ v1.34: ตัด orphan auto-salary records (otType="เงินเดือน") ออกจาก pending
    //   — กันบวกซ้ำกับ empData.salary (เกิดเมื่อ undo รอบจ่ายเก่าที่มี salary records)
    const pending = all.filter(r => !r.paidAt && r.name && r.date && r.otType !== "เงินเดือน" && isOnOrBeforeCutoff(r.date, cutoff));
    const carry   = all.filter(r => !r.paidAt && r.name && r.date && r.otType !== "เงินเดือน" && !isOnOrBeforeCutoff(r.date, cutoff));

    // group by employee + ดึง travelAllowance/socialSecurity จาก Employees
    const employees = await getEmployees(sheets);
    const empMap = {};
    employees.forEach(e => empMap[e.name] = e);

    const byEmp = {};
    pending.forEach(r => {
      byEmp[r.name] = byEmp[r.name] || { name: r.name, days: new Set(), hours: 0, holidays: 0, pay: 0, count: 0 };
      byEmp[r.name].days.add(r.date);
      byEmp[r.name].count += 1;
      byEmp[r.name].pay   += r.pay;
      if (r.otType === "วันธรรมดา") byEmp[r.name].hours += r.hours;
      else                          byEmp[r.name].holidays += 1;
    });
    // ★ v1.25: เพิ่ม travelAllowance + หัก socialSecurity (รายเดือน — รวมในรอบจ่าย)
    // ★ v1.30 BUG-05: ถ้าจ่ายไปแล้วในเดือนเดียวกัน → travel=0, social=0
    // ★ v1.33: salary ก็เป็น "monthly extras" เหมือน travel/social
    const empsAlreadyPaid = await getEmployeesAlreadyPaidExtrasThisMonth(sheets, all, cutoff);

    // ★ v1.33: เพิ่ม employees ที่มี salary > 0 แต่ไม่มี OT records ใน round นี้
    //   (เช่น เพิ่งเพิ่มเข้าระบบ เดือนนี้ไม่มี OT แต่ยังต้องได้เงินเดือน)
    employees.forEach(e => {
      if ((e.salary || 0) > 0 && !byEmp[e.name] && !empsAlreadyPaid.has(e.name)) {
        byEmp[e.name] = { name: e.name, days: new Set(), hours: 0, holidays: 0, pay: 0, count: 0 };
      }
    });

    const summary = Object.values(byEmp).map(e => {
      const empData = empMap[e.name] || {};
      const skipExtras = empsAlreadyPaid.has(e.name);
      const travel = skipExtras ? 0 : (empData.travelAllowance || 0);
      const social = skipExtras ? 0 : (empData.socialSecurity || 0);
      const salary = skipExtras ? 0 : (empData.salary || 0);     // ★ v1.33
      const netPay = e.pay + travel + salary - social;
      return {
        name: e.name, days: e.days.size, hours: +e.hours.toFixed(2),
        holidays: e.holidays, pay: e.pay, count: e.count,
        travel, social, salary, netPay,
        extrasSkipped: skipExtras,
      };
    }).sort((a,b) => b.netPay - a.netPay);

    const totals = {
      records: pending.length,
      employees: summary.length,
      totalPay: summary.reduce((s,e) => s+e.pay, 0),
      totalTravel: summary.reduce((s,e) => s+e.travel, 0),
      totalSocial: summary.reduce((s,e) => s+e.social, 0),
      totalSalary: summary.reduce((s,e) => s+(e.salary||0), 0),  // ★ v1.33
      totalNet: summary.reduce((s,e) => s+e.netPay, 0),
      totalHours: +pending.filter(r=>r.otType==="วันธรรมดา").reduce((s,r)=>s+r.hours,0).toFixed(2),
      totalHolidays: pending.filter(r=>r.otType!=="วันธรรมดา").length,
    };

    res.json({
      cutoff, totals, summary,
      records: pending,
      carry: carry.map(r => ({ name:r.name, date:r.date, startTime:r.startTime, endTime:r.endTime, hours:r.hours, otType:r.otType, pay:r.pay })),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/payroll/commit — ★ v1.32: เปลี่ยนเป็น "submit for approval" ────
//   - mark records กับ payId แล้ว (lock)
//   - log status = "pending_approval" (รอ supervisor อนุมัติ)
//   - finalize ทำใน /api/payroll/finalize (admin คนเดิม)
app.post("/api/payroll/commit", requireAdmin, async (req, res) => {
  const { cutoff, createdBy } = req.body;
  if (!cutoff || !/^\d{2}\/\d{2}\/\d{4}$/.test(cutoff)) {
    return res.status(400).json({ error: "cutoff ต้องอยู่ในรูป DD/MM/YYYY" });
  }
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    // ★ v1.34: ตัด orphan auto-salary records ออก (กันบวกซ้ำกับ empData.salary)
    const pending = all.filter(r => !r.paidAt && r.name && r.date && r.otType !== "เงินเดือน" && isOnOrBeforeCutoff(r.date, cutoff));

    if (pending.length === 0) return res.status(400).json({ error: "ไม่มีรายการให้จ่ายในรอบนี้" });

    // สร้าง payroll ID — PAY-YYYYMMDD-HHMM
    const now = new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Bangkok"}));
    const yyyy = now.getFullYear()+543;
    const mm   = String(now.getMonth()+1).padStart(2,"0");
    const dd   = String(now.getDate()).padStart(2,"0");
    const hh   = String(now.getHours()).padStart(2,"0");
    const min  = String(now.getMinutes()).padStart(2,"0");
    const payId = `PAY-${yyyy}${mm}${dd}-${hh}${min}`;
    const nowStr = `${dd}/${mm}/${yyyy} ${hh}:${min}`;

    // Mark column K ของแต่ละ pending record (batch update)
    const updates = pending.map(r => ({
      range: `OT_Records!K${idxToRow(r.idx)}`,
      values: [[payId]],
    }));

    await sheets.spreadsheets.values.batchUpdate({
      spreadsheetId: SHEET_ID,
      resource: {
        valueInputOption: "USER_ENTERED",
        data: updates,
      },
    });

    // คำนวณยอด (รวม travel allowance + เงินเดือน + หัก social security)
    const empList = await getEmployees(sheets);
    const empsAlreadyPaid = await getEmployeesAlreadyPaidExtrasThisMonth(sheets, all, cutoff);

    // ★ v1.33: เพิ่ม "พนักงานที่มี salary > 0 แต่ไม่มี OT รอบนี้" ที่ยังไม่ได้จ่าย salary เดือนนี้
    //   → ใส่ records "เงินเดือน" ลง OT_Records ก่อน (paidAt = payId) เพื่อให้ dedup logic ครอบคลุม
    const otEmployeeNames = new Set(pending.map(r => r.name));
    const salaryRowsToAppend = [];
    for (const e of empList) {
      if ((e.salary || 0) <= 0) continue;
      if (empsAlreadyPaid.has(e.name)) continue;
      if (otEmployeeNames.has(e.name)) continue;  // มี OT แล้ว — รอบเดียวกันบวก salary ผ่าน totalSalary
      // เพิ่ม dummy record "เงินเดือน" เพื่อ track การจ่าย
      salaryRowsToAppend.push([
        e.name, cutoff, "-", "-", 0, "เงินเดือน (auto)", "",
        "เงินเดือน", e.salary, new Date().toLocaleString("th-TH",{timeZone:"Asia/Bangkok"}),
      ]);
    }

    const employeeNames = new Set(pending.map(r => r.name));
    salaryRowsToAppend.forEach((_, i) => employeeNames.add(salaryRowsToAppend[i][0]));
    const grossOT  = pending.reduce((s,r) => s+r.pay, 0);

    let totalTravel = 0, totalSocial = 0, totalSalary = 0;
    let extrasSkippedCount = 0;
    for (const name of employeeNames) {
      if (empsAlreadyPaid.has(name)) { extrasSkippedCount++; continue; }
      const e = empList.find(x => x.name === name);
      if (e) {
        totalTravel += e.travelAllowance || 0;
        totalSocial += e.socialSecurity || 0;
        totalSalary += e.salary || 0;
      }
    }
    if (extrasSkippedCount > 0) {
      console.log(`💡 Skip travel/social/salary สำหรับ ${extrasSkippedCount} คน (จ่ายไปแล้วในเดือนนี้)`);
    }
    const totalPay = grossOT + totalTravel + totalSalary - totalSocial;  // net

    // ★ v1.33: append "salary records" สำหรับพนักงานที่ไม่มี OT รอบนี้แต่ต้องรับเงินเดือน
    //   ใส่ paidAt (col K) ตรง ๆ → dedup logic เห็นทันที
    if (salaryRowsToAppend.length > 0) {
      // เพิ่ม payId เป็น col K
      const rowsWithPayId = salaryRowsToAppend.map(r => [...r, payId]);
      try {
        await sheets.spreadsheets.values.append({
          spreadsheetId: SHEET_ID,
          range: "OT_Records!A:K",
          valueInputOption: "USER_ENTERED",
          resource: { values: rowsWithPayId },
        });
        console.log(`💼 เพิ่ม salary records ${salaryRowsToAppend.length} แถว`);
      } catch (e) {
        console.error("Append salary records failed:", e.message);
      }
    }

    // เขียน Payroll_Log (ถ้า tab มีอยู่)
    // ★ v1.32: status = "pending_approval" (เดิม "active")
    const totalRecords = pending.length + salaryRowsToAppend.length;  // ★ v1.33
    try {
      await sheets.spreadsheets.values.append({
        spreadsheetId: SHEET_ID,
        range: "Payroll_Log!A:J",
        valueInputOption: "USER_ENTERED",
        resource: { values: [[
          payId, cutoff, nowStr,
          totalRecords, employeeNames.size, totalPay,
          createdBy || "Admin", "pending_approval", "", "",
        ]] },
      });
    } catch (logErr) {
      console.error("Payroll_Log write failed (tab อาจยังไม่มี):", logErr.message);
    }

    res.json({
      ok: true,
      payId,
      cutoff,
      submittedAt: nowStr,
      status: "pending_approval",
      records: totalRecords,        // ★ v1.33
      employees: employeeNames.size,
      grossOT,
      totalTravel,
      totalSocial,
      totalSalary,                  // ★ v1.33
      totalPay,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/payroll/undo — ปลด record ในรอบนั้นกลับเป็น pending ──
// ★ v1.17.4: soft cleanup — ถ้าไม่เจอ record ก็ mark log เป็น undone (orphan cleanup)
app.post("/api/payroll/undo", requireAdmin, async (req, res) => {
  const { payId, undoneBy } = req.body;
  if (!payId) return res.status(400).json({ error: "ระบุ payId" });
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    const target = all.filter(r => r.paidAt === payId);

    // ★ v1.34: แยก records 2 ประเภท
    //   - "เงินเดือน" auto-records → DELETE ทิ้ง (ไม่ใช่ OT จริง สร้างมาเพื่อ track salary)
    //   - records OT/holiday จริง   → ปลด column K (paidAt) คืน เป็น pending
    const autoSalaryRecs = target.filter(r => r.otType === "เงินเดือน");
    const realOtRecs     = target.filter(r => r.otType !== "เงินเดือน");

    // ปลด column K ของ OT records จริง
    if (realOtRecs.length > 0) {
      const updates = realOtRecs.map(r => ({
        range: `OT_Records!K${idxToRow(r.idx)}`,
        values: [[""]],
      }));
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        resource: { valueInputOption: "USER_ENTERED", data: updates },
      });
    }

    // ลบ auto-salary records ทิ้ง — ลบจาก row สูงไปต่ำ (กัน index shift)
    if (autoSalaryRecs.length > 0) {
      const sortedDesc = [...autoSalaryRecs].sort((a, b) => b.idx - a.idx);
      for (const r of sortedDesc) {
        try {
          await deleteRow(sheets, "OT_Records", idxToRow(r.idx));
        } catch (delErr) {
          console.error(`ลบ salary record row ${idxToRow(r.idx)} ไม่สำเร็จ:`, delErr.message);
        }
      }
      console.log(`🗑️  ลบ auto-salary records ${autoSalaryRecs.length} แถว (payId=${payId})`);
    }

    // ★ v1.17.5: อัปเดต Payroll_Log status = "undone" — ทุก row ที่ match payId (กันมี duplicate)
    let logUpdated = 0;
    try {
      const r = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: "Payroll_Log!A3:J5000",
      });
      const rows = r.data.values || [];
      const matches = [];
      rows.forEach((row, i) => {
        if ((row[0] || "").trim() === payId && (row[7] || "active").trim() !== "undone") {
          matches.push(i);
        }
      });
      if (matches.length > 0) {
        const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
        const updates = matches.map(idx => ({
          range: `Payroll_Log!H${idx + 3}:J${idx + 3}`,
          values: [["undone", now, undoneBy || "Admin"]],
        }));
        await sheets.spreadsheets.values.batchUpdate({
          spreadsheetId: SHEET_ID,
          resource: { valueInputOption: "USER_ENTERED", data: updates },
        });
        logUpdated = matches.length;
      }
    } catch (logErr) {
      console.error("Payroll_Log update failed:", logErr.message);
    }

    // ถ้าไม่เจอทั้ง records และ log → orphan ที่หาไม่เจอ
    if (target.length === 0 && logUpdated === 0) {
      return res.status(400).json({ error: `ไม่พบรอบจ่าย ${payId} ทั้งใน records และ Payroll_Log (อาจถูกยกเลิกไปแล้ว)` });
    }

    res.json({
      ok: true,
      payId,
      recordsRestored: realOtRecs.length,           // ★ v1.34
      autoSalaryDeleted: autoSalaryRecs.length,     // ★ v1.34
      logCleanedOnly: target.length === 0,
      logRowsUpdated: logUpdated,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// ★ v1.34: CLEANUP ORPHAN AUTO-SALARY RECORDS
//   หลังจาก undo รอบจ่ายเก่า — auto-salary records (otType="เงินเดือน")
//   จะค้างเป็น pending ทำให้ preview/commit นับ salary ซ้ำ
//   endpoint นี้สแกนหา records พวกนี้แล้วลบทิ้ง (ไม่ใช่ OT จริง)
// ══════════════════════════════════════════════════════════════
app.get("/api/payroll/cleanup-orphan-salary/preview", requireAdmin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const all = await getAllRecords(sheets);
    const orphans = all.filter(r => r.otType === "เงินเดือน" && !r.paidAt && r.name);
    res.json({
      ok: true,
      count: orphans.length,
      totalAmount: orphans.reduce((s, r) => s + (r.pay || 0), 0),
      items: orphans.map(r => ({
        idx: r.idx, row: idxToRow(r.idx),
        name: r.name, date: r.date, pay: r.pay,
      })),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

app.post("/api/payroll/cleanup-orphan-salary/apply", requireAdmin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const all = await getAllRecords(sheets);
    const orphans = all.filter(r => r.otType === "เงินเดือน" && !r.paidAt && r.name);

    if (orphans.length === 0) {
      return res.json({ ok: true, deleted: 0, message: "ไม่มี records ค้าง" });
    }

    // ลบจาก row สูงไปต่ำ (กัน index shift)
    const sortedDesc = [...orphans].sort((a, b) => b.idx - a.idx);
    let deleted = 0;
    for (const r of sortedDesc) {
      try {
        await deleteRow(sheets, "OT_Records", idxToRow(r.idx));
        deleted++;
      } catch (delErr) {
        console.error(`ลบ row ${idxToRow(r.idx)} ไม่สำเร็จ:`, delErr.message);
      }
    }
    console.log(`🗑️  Cleanup orphan auto-salary: ลบ ${deleted}/${orphans.length} แถว`);
    res.json({
      ok: true,
      deleted,
      totalAmount: orphans.slice(0, deleted).reduce((s, r) => s + (r.pay || 0), 0),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// ★ v1.34: RECALC PENDING — คำนวณใหม่ records ที่ยังไม่จ่าย
//   ใช้เวลาเปลี่ยนสูตร (เช่น เปลี่ยน WEEKDAY_MULTIPLIER)
//   เพื่อปรับ pay ของ records ที่ค้างอยู่ให้ตรงกับสูตรปัจจุบัน
// ══════════════════════════════════════════════════════════════

// helper: คำนวณ pay ที่ "ควรจะเป็น" สำหรับ record หนึ่งตัว ตามสูตรปัจจุบัน
function expectedPay(rec, emp) {
  if (!emp) return null;
  if (rec.otType === "วันธรรมดา") {
    return Math.round((rec.hours || 0) * (emp.hourlyRate || 0) * WEEKDAY_MULTIPLIER);
  }
  // วันหยุด/อาทิตย์/ตจว. — ใช้ flat rate
  if (rec.otType && rec.otType.includes("ต่างจังหวัด")) {
    return Math.round(emp.outProvinceFlat || 0);
  }
  // เงินเดือน auto record — ไม่แตะ
  if (rec.otType === "เงินเดือน") return rec.pay;
  return Math.round(emp.holidayFlat || 0);
}

// ── GET /api/payroll/recalc-preview — ดู diff ก่อน apply ──
app.get("/api/payroll/recalc-preview", requireAdmin, async (req, res) => {
  try {
    const sheets    = await getSheetsClient();
    const all       = await getAllRecords(sheets);
    const employees = await getEmployees(sheets);
    const empMap    = Object.fromEntries(employees.map(e => [e.name, e]));

    const diffs = [];
    for (const r of all) {
      if (r.paidAt) continue;             // ข้าม records ที่จ่ายแล้ว
      if (!r.name)  continue;
      if (r.otType === "เงินเดือน") continue; // ข้าม salary auto records
      const emp = empMap[r.name];
      if (!emp) continue;
      const newPay = expectedPay(r, emp);
      if (newPay === null) continue;
      if (newPay !== r.pay) {
        diffs.push({
          idx: r.idx,
          row: idxToRow(r.idx),
          name: r.name,
          date: r.date,
          startTime: r.startTime,
          endTime: r.endTime,
          hours: r.hours,
          otType: r.otType,
          oldPay: r.pay,
          newPay,
          diff: newPay - r.pay,
        });
      }
    }

    res.json({
      ok: true,
      formula: { weekdayMultiplier: WEEKDAY_MULTIPLIER },
      count: diffs.length,
      totalOldPay: diffs.reduce((s, d) => s + d.oldPay, 0),
      totalNewPay: diffs.reduce((s, d) => s + d.newPay, 0),
      totalDiff:   diffs.reduce((s, d) => s + d.diff,   0),
      items: diffs,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/payroll/recalc-apply — ลงมือแก้ pay ในชีต ──
app.post("/api/payroll/recalc-apply", requireAdmin, async (req, res) => {
  try {
    const sheets    = await getSheetsClient();
    const all       = await getAllRecords(sheets);
    const employees = await getEmployees(sheets);
    const empMap    = Object.fromEntries(employees.map(e => [e.name, e]));

    const updates = [];
    const applied = [];
    for (const r of all) {
      if (r.paidAt) continue;
      if (!r.name)  continue;
      if (r.otType === "เงินเดือน") continue;
      const emp = empMap[r.name];
      if (!emp) continue;
      const newPay = expectedPay(r, emp);
      if (newPay === null) continue;
      if (newPay !== r.pay) {
        const row = idxToRow(r.idx);
        updates.push({
          range: `OT_Records!I${row}`,    // col I = pay
          values: [[newPay]],
        });
        applied.push({ name: r.name, date: r.date, oldPay: r.pay, newPay, row });
      }
    }

    if (updates.length === 0) {
      return res.json({ ok: true, count: 0, message: "ไม่มี records ที่ต้องแก้" });
    }

    // batch update เป็น chunk ละ 100 row กัน timeout
    for (let i = 0; i < updates.length; i += 100) {
      const chunk = updates.slice(i, i + 100);
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        resource: { valueInputOption: "USER_ENTERED", data: chunk },
      });
    }

    console.log(`🔧 Recalc applied: ${applied.length} records updated`);
    res.json({
      ok: true,
      count: applied.length,
      totalDiff: applied.reduce((s, a) => s + (a.newPay - a.oldPay), 0),
      applied,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// ★ v1.32: APPROVAL FLOW ENDPOINTS
//   - /approve/:payId   : supervisor → status pending_approval → approved
//   - /reject/:payId    : supervisor → ปลด records, status → rejected
//   - /finalize/:payId  : admin → status approved → active (จบ flow)
// ══════════════════════════════════════════════════════════════

// ★ v1.32 dual-approval helpers ─────────────────────────
// list ของชื่อ supervisor ทั้งหมด (role==="supervisor" ใน Employees)
async function getRequiredApprovers(sheets) {
  const employees = await getEmployees(sheets);
  return employees.filter(e => e.role === "supervisor").map(e => e.name);
}
// parse "Bow,Woot" → ["Bow","Woot"]
function parseApprovers(str) {
  return (str || "").split(",").map(s => s.trim()).filter(Boolean);
}
// resolve userId → employee name (สำหรับเก็บ approvers ลง sheet)
async function getNameByUserId(sheets, uid) {
  const employees = await getEmployees(sheets);
  const e = employees.find(x => x.userId === uid);
  return e ? e.name : "Unknown";
}

// helper: หา row ของ payId ใน Payroll_Log
async function findPayrollLogRow(sheets, payId) {
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Payroll_Log!A3:J5000",
  });
  const rows = r.data.values || [];
  for (let i = 0; i < rows.length; i++) {
    if ((rows[i][0] || "").trim() === payId) {
      return { idx: i, row: rows[i], rowNum: i + 3 };  // rowNum = sheet row (1-based, header offset)
    }
  }
  return null;
}

// ── POST /api/payroll/approve/:payId — Supervisor อนุมัติ (dual-approval)
//   ★ v1.32: ต้องมี supervisor ครบทุกคน (Bow + Woot ฯลฯ) ถึงจะ flip → approved
app.post("/api/payroll/approve/:payId", requireSupervisor, async (req, res) => {
  const { payId } = req.params;
  try {
    const sheets = await getSheetsClient();
    const found = await findPayrollLogRow(sheets, payId);
    if (!found) return res.status(404).json({ error: `ไม่พบรอบ ${payId}` });
    const status = (found.row[7] || "").trim();
    if (status !== "pending_approval") {
      return res.status(400).json({ error: `รอบนี้สถานะ "${status}" — อนุมัติไม่ได้` });
    }
    const uid = (req.headers["x-line-user-id"] || "").toString().trim();
    const callerName = await getNameByUserId(sheets, uid);
    const required   = await getRequiredApprovers(sheets);  // ["Bow","Woot"]
    const current    = new Set(parseApprovers(found.row[9])); // col J → existing approvers

    if (current.has(callerName)) {
      return res.status(400).json({ error: `${callerName} อนุมัติรอบนี้ไปแล้ว` });
    }
    current.add(callerName);
    const approversStr = [...current].join(",");

    // ครบหรือยัง? (caller ที่ไม่ใช่ supervisor — ถือว่าช่วย approve แต่ไม่นับใน required)
    const approvedRequired = required.filter(n => current.has(n));
    const allRequiredDone  = required.length > 0 && approvedRequired.length === required.length;

    const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
    const newStatus = allRequiredDone ? "approved" : "pending_approval";
    // เขียน H (status), I (lastActionAt), J (approvers comma-sep)
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Payroll_Log!H${found.rowNum}:J${found.rowNum}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[newStatus, now, approversStr]] },
    });
    res.json({
      ok: true, payId, status: newStatus,
      approvedAt: now, approvedBy: callerName,
      approvers: [...current],
      required, remaining: required.filter(n => !current.has(n)),
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/payroll/reject/:payId — Supervisor ไม่อนุมัติ → ปลด records กลับ unpaid
app.post("/api/payroll/reject/:payId", requireSupervisor, async (req, res) => {
  const { payId } = req.params;
  const { rejectedBy, reason } = req.body || {};
  try {
    const sheets = await getSheetsClient();
    const found = await findPayrollLogRow(sheets, payId);
    if (!found) return res.status(404).json({ error: `ไม่พบรอบ ${payId}` });
    const status = (found.row[7] || "").trim();
    if (status !== "pending_approval") {
      return res.status(400).json({ error: `รอบนี้สถานะ "${status}" — ไม่อนุมัติได้เฉพาะ pending_approval` });
    }

    // ปลด records (clear column K — paidAt)
    const all = await getAllRecords(sheets);
    const target = all.filter(r => r.paidAt === payId);
    if (target.length > 0) {
      const updates = target.map(r => ({
        range: `OT_Records!K${idxToRow(r.idx)}`,
        values: [[""]],
      }));
      await sheets.spreadsheets.values.batchUpdate({
        spreadsheetId: SHEET_ID,
        resource: { valueInputOption: "USER_ENTERED", data: updates },
      });
    }

    const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
    const note = (rejectedBy || "Supervisor") + (reason ? ` (${reason})` : "");
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Payroll_Log!H${found.rowNum}:J${found.rowNum}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [["rejected", now, note]] },
    });
    res.json({ ok: true, payId, status: "rejected", recordsRestored: target.length, rejectedAt: now });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/payroll/finalize/:payId — Admin บันทึก+download หลัง approved
app.post("/api/payroll/finalize/:payId", requireAdmin, async (req, res) => {
  const { payId } = req.params;
  const { finalizedBy } = req.body || {};
  try {
    const sheets = await getSheetsClient();
    const found = await findPayrollLogRow(sheets, payId);
    if (!found) return res.status(404).json({ error: `ไม่พบรอบ ${payId}` });
    const status = (found.row[7] || "").trim();
    if (status !== "approved") {
      return res.status(400).json({ error: `รอบนี้สถานะ "${status}" — ต้อง approved ก่อนถึง finalize ได้` });
    }
    const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Payroll_Log!H${found.rowNum}:J${found.rowNum}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [["active", now, finalizedBy || "Admin"]] },
    });
    res.json({ ok: true, payId, status: "active", finalizedAt: now });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/payroll/pending — Supervisor: list รอบ pending_approval + records
app.get("/api/payroll/pending", requireSupervisor, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID, range: "Payroll_Log!A3:J5000",
    });
    const logRows = (r.data.values || []).map((row, i) => ({
      idx: i, payId: (row[0]||"").trim(), cutoff: (row[1]||"").trim(),
      createdAt: row[2] || "", records: Number(row[3]) || 0,
      employees: Number(row[4]) || 0, totalPay: Number(row[5]) || 0,
      createdBy: row[6] || "", status: (row[7]||"").trim(),
      approvers: parseApprovers(row[9]),  // ★ v1.32: approvers list จาก col J
    })).filter(p => p.payId && p.status === "pending_approval");

    // ★ v1.32: required approvers + caller — frontend ใช้ตัดสิน UI
    const required = await getRequiredApprovers(sheets);
    const callerUid = (req.headers["x-line-user-id"] || "").toString().trim();
    const callerName = await getNameByUserId(sheets, callerUid);

    // attach summary per payId
    const all = await getAllRecords(sheets);
    const employees = await getEmployees(sheets);
    const empMap = {};
    employees.forEach(e => empMap[e.name] = e);

    const items = logRows.map(p => {
      const recs = all.filter(r => r.paidAt === p.payId);
      const byEmp = {};
      recs.forEach(r => {
        byEmp[r.name] = byEmp[r.name] || { name: r.name, days: new Set(), hours: 0, holidays: 0, pay: 0, count: 0 };
        byEmp[r.name].days.add(r.date);
        byEmp[r.name].count += 1;
        byEmp[r.name].pay   += r.pay;
        if (r.otType === "วันธรรมดา") byEmp[r.name].hours += r.hours;
        else                          byEmp[r.name].holidays += 1;
      });
      const summary = Object.values(byEmp).map(e => {
        const ed = empMap[e.name] || {};
        return {
          name: e.name, days: e.days.size, hours: +e.hours.toFixed(2),
          holidays: e.holidays, pay: e.pay, count: e.count,
          travel: ed.travelAllowance || 0, social: ed.socialSecurity || 0,
        };
      }).sort((a,b) => b.pay - a.pay);
      return { ...p, summary };
    }).reverse(); // ใหม่บนสุด

    res.json({ items, required, callerName });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/payroll/history — ดูประวัติการจ่ายเงิน ──────────
app.get("/api/payroll/history", requireAdmin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Payroll_Log!A3:J5000",
    });
    // ★ v1.32: required approvers (สำหรับแสดง progress ในรายการ pending)
    const required = await getRequiredApprovers(sheets);
    const rows = (r.data.values || []).map((row, i) => {
      const status = row[7] || "active";
      const colJ = row[9] || "";  // approvers สำหรับ pending/approved, ชื่อสำหรับ rejected/undone
      return {
        idx: i,
        payId: row[0] || "",
        cutoff: row[1] || "",
        createdAt: row[2] || "",
        records: Number(row[3]) || 0,
        employees: Number(row[4]) || 0,
        totalPay: Number(row[5]) || 0,
        createdBy: row[6] || "",
        status,
        undoneAt: row[8] || "",        // = lastActionAt
        undoneBy: status === "undone" || status === "rejected" ? colJ : "",
        approvers: (status === "pending_approval" || status === "approved") ? parseApprovers(colJ) : [],
      };
    }).filter(r => r.payId);
    res.json({ history: rows.reverse(), required });  // ★ ส่ง required ด้วย
  } catch (e) {
    res.json({ history: [], required: [] });
  }
});

// ── DELETE /api/admin/records/:idx — ★ v1.34: Admin force-delete OT record (ฉุกเฉิน)
//   ใช้กรณีพนักงาน login เข้าระบบไม่ได้ → admin ลบให้แทน
//   เงื่อนไข: record ต้องยังไม่จ่าย (paidAt ว่าง) — ถ้าจ่ายแล้วต้อง undo รอบก่อน
app.delete("/api/admin/records/:idx", requireAdmin, async (req, res) => {
  const idx = Number(req.params.idx);
  if (!Number.isInteger(idx) || idx < 0) {
    return res.status(400).json({ error: "idx ไม่ถูกต้อง" });
  }
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    const target = all.find(r => r.idx === idx);
    if (!target) return res.status(404).json({ error: `ไม่พบ record idx=${idx}` });
    if (target.paidAt) {
      return res.status(400).json({ error: `record นี้จ่ายไปแล้ว (รอบ ${target.paidAt}) ลบไม่ได้ — ต้อง undo รอบจ่ายก่อน` });
    }
    await deleteRow(sheets, "OT_Records", idxToRow(idx));
    const adminName = req.headers["x-admin-name"] || "Admin";
    console.log(`🗑️  Admin force-delete: ${target.name} ${target.date} (idx=${idx}, by ${adminName})`);
    res.json({
      ok: true,
      idx,
      name: target.name,
      date: target.date,
      otType: target.otType,
      pay: target.pay,
    });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/admin/records — records สำหรับ Admin + Supervisor (ดูภาพรวม)
// ★ v1.32 fix: เปลี่ยนเป็น requireSupervisor — supervisor ต้องเห็นข้อมูลภาพรวม + กราฟ
// ★ v1.34: + pendingCarry → records pending จากเดือนก่อน (ที่ตกค้างยกยอดมา)
app.get("/api/admin/records", requireSupervisor, async (req, res) => {
  const { month, year } = req.query;
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    // records ของเดือนที่เลือก
    const rows   = all.filter(r =>
      !month || r.date.includes(`/${month}/${year}`)
    );

    // ★ v1.34: pendingCarry = records pending ที่อยู่ก่อนเดือน month/year
    //   (ตกค้างจ่ายข้ามเดือน — เกิดจากลืมลง / ลงไม่ทัน / cutoff ก่อนสิ้นเดือน)
    let pendingCarry = [];
    if (month && year) {
      const m = Number(month), y = Number(year);
      // แปลงเป็น timestamp ของวันที่ 1 เดือน month/year (พ.ศ. → ค.ศ.)
      const monthStart = new Date(y - 543, m - 1, 1).getTime();
      pendingCarry = all.filter(r => {
        if (r.paidAt) return false;
        if (r.otType === "เงินเดือน") return false;
        if (!r.date) return false;
        const [d, mm2, yy2] = r.date.split("/").map(Number);
        if (!d || !mm2 || !yy2) return false;
        const t = new Date(yy2 - 543, mm2 - 1, d).getTime();
        return t < monthStart;
      });
    }

    res.json({ records: rows, pendingCarry });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/edit-requests — คำขอแก้ไข (Admin) ──────────────
app.get("/api/edit-requests", requireAdmin, async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    const r      = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Edit_Requests!A3:F5000",  // ★ v1.14: ขยายจาก 500 → 5000 (~12 ปี)
    });
    const rows = (r.data.values || []).map((row, i) => ({
      idx: i, name: row[0], date: row[1], recordDesc: row[2],
      note: row[3], status: row[4] || "รอดำเนินการ", createdAt: row[5],
    }));
    res.json({ requests: rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── PUT /api/edit-requests/:idx — อนุมัติ/ปฏิเสธ (Admin) ─────
app.put("/api/edit-requests/:idx", requireAdmin, async (req, res) => {
  const row    = idxToRow(Number(req.params.idx));
  const { status } = req.body;
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Edit_Requests!E${row}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[status]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ★ v1.22: POST /api/edit-requests/:idx/apply — อนุมัติ + แก้ record จริง
// ★ v1.22.1: รองรับทั้ง weekday + holiday + delete record
app.post("/api/edit-requests/:idx/apply", requireAdmin, async (req, res) => {
  const reqIdx = Number(req.params.idx);
  const { action, newStartTime, newEndTime, newTask, newLocation, newDate } = req.body;
  // action: "edit" (default) | "delete"
  try {
    const sheets = await getSheetsClient();

    // 1. Get the edit request
    const reqRes = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Edit_Requests!A3:F5000",
    });
    const reqRows = reqRes.data.values || [];
    if (reqIdx >= reqRows.length) return res.status(400).json({ error: "ไม่พบคำขอ" });
    const reqRow = reqRows[reqIdx];
    const reqName = (reqRow[0] || "").trim();
    const reqDate = (reqRow[1] || "").trim();
    const reqDesc = reqRow[2] || "";
    const reqStatus = reqRow[4] || "รอดำเนินการ";
    if (reqStatus !== "รอดำเนินการ") {
      return res.status(400).json({ error: `คำขอนี้สถานะ "${reqStatus}" แล้ว — ดำเนินการไปแล้ว` });
    }

    // 2. Find target record
    const all = await getAllRecords(sheets);
    const candidates = all.filter(r => r.name === reqName && r.date === reqDate);
    if (candidates.length === 0) {
      return res.status(400).json({ error: `ไม่พบ record ของ ${reqName} วันที่ ${reqDate}` });
    }
    // Match by old startTime in recordDesc (if multiple weekday)
    const timeMatch = reqDesc.match(/(\d{2}:\d{2})/);
    let target = (timeMatch && candidates.length > 1)
      ? candidates.find(r => r.startTime === timeMatch[1])
      : null;
    target = target || candidates[0];

    if (target.paidAt) {
      return res.status(400).json({ error: `รายการนี้จ่ายไปแล้ว (รอบ ${target.paidAt}) แก้ไม่ได้` });
    }

    const isWeekday = target.otType === "วันธรรมดา";

    // ── DELETE action ──
    if (action === "delete") {
      const targetRow = idxToRow(target.idx);
      await deleteRow(sheets, "OT_Records", targetRow);
      // Mark Edit_Request as approved
      // หลัง delete row → idx ของ Edit_Requests อาจไม่ถูกเปลี่ยน (เป็นคนละ tab)
      const reqRowSheet = idxToRow(reqIdx);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Edit_Requests!E${reqRowSheet}`,
        valueInputOption: "USER_ENTERED",
        resource: { values: [["อนุมัติแล้ว (ลบ)"]] },
      });
      return res.json({ ok: true, deleted: true });
    }

    // ── EDIT action (default) ──
    if (isWeekday) {
      if (!newStartTime || !newEndTime || !newTask) {
        return res.status(400).json({ error: "กรุณากรอก เริ่ม, สิ้นสุด, งาน" });
      }
      // ★ v1.27: ใช้ window per-employee (เสาร์ครึ่งวัน)
      const employeesForCheck = await getEmployees(sheets);
      const empForCheck = employeesForCheck.find(e => e.name === reqName);
      const editDate = newDate || target.date;
      const [eDD, eMM, eYY] = editDate.split("/").map(Number);
      const eDow = new Date(eYY - 543, eMM - 1, eDD).getDay();
      const editWorkEnd = getWorkEndForEmp(empForCheck, eDow);
      // Validate
      if (overlapsWorkHours(newStartTime, newEndTime, editWorkEnd)) {
        return res.status(400).json({ error: `ช่วงเวลาใหม่ทับเวลางานปกติ ${workWindowLabel(editWorkEnd)}` });
      }
      const newHours = calcHours(newStartTime, newEndTime);
      if (newHours <= 0) return res.status(400).json({ error: "เวลาสิ้นสุดต้องมากกว่าเริ่มต้น" });

      // Check overlap with OTHER same-day records (skip target itself)
      const others = all.filter(r =>
        r.name === reqName && r.date === reqDate &&
        r.otType === "วันธรรมดา" && r.idx !== target.idx &&
        r.startTime !== "-" && r.endTime !== "-"
      );
      const toMin = t => { const [h,m] = t.split(":").map(Number); return h*60+m; };
      let nS = toMin(newStartTime), nE = toMin(newEndTime);
      if (nE < nS) nE += 1440;
      for (const r of others) {
        let s = toMin(r.startTime), e = toMin(r.endTime);
        if (e < s) e += 1440;
        if (nS < e && s < nE) {
          return res.status(400).json({ error: `ช่วงใหม่ ${newStartTime}-${newEndTime} ทับกับรายการอื่น ${r.startTime}-${r.endTime}` });
        }
      }

      // Recalc pay
      const employees = await getEmployees(sheets);
      const emp = employees.find(e => e.name === reqName);
      if (!emp) return res.status(400).json({ error: `ไม่พบพนักงาน ${reqName}` });
      const otherHours = others.reduce((s, r) => s + r.hours, 0);
      const remainingPayable = Math.max(0, MAX_OT_PER_DAY - otherHours);
      const payableHours = Math.min(newHours, remainingPayable);
      const newPay = Math.round(payableHours * emp.hourlyRate * WEEKDAY_MULTIPLIER);

      // Update OT_Records row (B-I: date, startTime, endTime, hours, task, location, otType, pay)
      const targetRow = idxToRow(target.idx);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `OT_Records!B${targetRow}:I${targetRow}`,
        valueInputOption: "USER_ENTERED",
        resource: { values: [[
          newDate || target.date,
          newStartTime, newEndTime, newHours,
          newTask, newLocation || "", "วันธรรมดา", newPay,
        ]] },
      });

      // Mark approved
      const reqRowSheet = idxToRow(reqIdx);
      await sheets.spreadsheets.values.update({
        spreadsheetId: SHEET_ID,
        range: `Edit_Requests!E${reqRowSheet}`,
        valueInputOption: "USER_ENTERED",
        resource: { values: [["อนุมัติแล้ว"]] },
      });

      return res.json({ ok: true, newHours, payableHours, newPay, capped: payableHours < newHours });
    }

    // ── HOLIDAY record edit (just task/location/date) ──
    if (!newTask) return res.status(400).json({ error: "กรุณากรอกงานที่ทำ" });

    const targetRow = idxToRow(target.idx);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `OT_Records!B${targetRow}:I${targetRow}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[
        newDate || target.date,
        "-", "-", 0,
        newTask, newLocation || "", target.otType, target.pay,
      ]] },
    });

    const reqRowSheet = idxToRow(reqIdx);
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Edit_Requests!E${reqRowSheet}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [["อนุมัติแล้ว"]] },
    });

    return res.json({ ok: true, holiday: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ══════════════════════════════════════════════════════════════
// SERVER-SIDE EXCEL EXPORT (v1.16) — สำหรับ LIFF mobile
// ══════════════════════════════════════════════════════════════

// ── GET /api/export/monthly?employee=&month=&year= ──────────
app.get("/api/export/monthly", requireAdmin, async (req, res) => {
  let { employee = "", month, year } = req.query;
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    const d = new Date();
    const mm = month || String(d.getMonth()+1).padStart(2,"0");
    const yy = year  || String(d.getFullYear()+543);

    const recs = all.filter(r => {
      if (!r.name || !r.date) return false;
      if (employee && r.name !== employee) return false;
      const parts = r.date.split("/");
      return parts[1] === mm && parts[2] === yy;
    });

    if (recs.length === 0) {
      return res.status(404).send("ไม่มีข้อมูล OT ในเดือนนี้");
    }

    const wb = XLSX.utils.book_new();

    // ── Summary ──
    const byEmp = {};
    recs.forEach(r => {
      byEmp[r.name] = byEmp[r.name] || { name: r.name, days: new Set(), hours: 0, holidays: 0, pay: 0, count: 0 };
      byEmp[r.name].days.add(r.date);
      byEmp[r.name].count += 1;
      byEmp[r.name].pay   += r.pay;
      if (r.otType === "วันธรรมดา") byEmp[r.name].hours += r.hours;
      else                          byEmp[r.name].holidays += 1;
    });
    // ★ v1.25: เพิ่ม travel allowance + social security
    const empListM = await getEmployees(sheets);
    const empMapM = {};
    empListM.forEach(e => empMapM[e.name] = e);
    const summary = Object.values(byEmp).map(e => {
      const ed = empMapM[e.name] || {};
      const travel = ed.travelAllowance || 0;
      const social = ed.socialSecurity || 0;
      const salary = ed.salary || 0;     // ★ v1.33
      const netPay = e.pay + travel + salary - social;
      return { name: e.name, days: e.days.size, hours: +e.hours.toFixed(2),
               holidays: e.holidays, pay: e.pay, count: e.count, travel, social, salary, netPay };
    }).sort((a,b) => b.netPay - a.netPay);

    const sumRows = [
      [`สรุป OT — เดือน ${mm}/${yy}` + (employee ? ` — ${employee}` : "")],
      [],
      ["ลำดับ","ชื่อพนักงาน","จำนวนวัน","ชั่วโมง","วันหยุด/ตจว.","ค่า OT (฿)","+ ค่าเดินทาง","+ เงินเดือน","- ประกันสังคม","สุทธิ (฿)"],
    ];
    summary.forEach((e, i) => sumRows.push([
      i+1, e.name, e.days, e.hours, e.holidays, e.pay, e.travel, e.salary, e.social, e.netPay,
    ]));
    sumRows.push([]);
    sumRows.push([
      "รวมทั้งหมด", "",
      summary.reduce((s,e)=>s+e.days,0),
      +summary.reduce((s,e)=>s+e.hours,0).toFixed(2),
      summary.reduce((s,e)=>s+e.holidays,0),
      summary.reduce((s,e)=>s+e.pay,0),
      summary.reduce((s,e)=>s+e.travel,0),
      summary.reduce((s,e)=>s+(e.salary||0),0),  // ★ v1.33
      summary.reduce((s,e)=>s+e.social,0),
      summary.reduce((s,e)=>s+e.netPay,0),
    ]);
    const ws1 = XLSX.utils.aoa_to_sheet(sumRows);
    ws1["!cols"] = [{wch:8},{wch:18},{wch:10},{wch:10},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];
    ws1["!merges"] = [{ s: { r: 0, c: 0 }, e: { r: 0, c: 9 } }];
    styleSheet(ws1, { titleRow: 0, headerRow: 2 });
    applyA4Print(ws1, { orientation: "landscape" });
    XLSX.utils.book_append_sheet(wb, ws1, "สรุป");
    setPrintTitles(wb, "สรุป", 2);

    // ── Details ──
    const detailRows = [
      ["ชื่อ","วันที่","เริ่ม","สิ้นสุด","ชม.","งาน","สถานที่","ประเภท","ค่า (฿)","สถานะ"],
    ];
    recs.slice().sort((a,b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date)).forEach(r => {
      detailRows.push([
        r.name, r.date, r.startTime, r.endTime, r.hours,
        r.task||"", r.location||"", r.otType, r.pay,
        r.paidAt ? "✅ จ่ายแล้ว" : "⏳ รอจ่าย",
      ]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(detailRows);
    ws2["!cols"] = [{wch:18},{wch:13},{wch:8},{wch:8},{wch:8},{wch:28},{wch:16},{wch:18},{wch:12},{wch:14}];
    styleSheet(ws2, { headerRow: 0 });
    applyA4Print(ws2, { orientation: "landscape" });
    XLSX.utils.book_append_sheet(wb, ws2, "รายละเอียด");
    setPrintTitles(wb, "รายละเอียด", 0);

    const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
    const fname = (employee ? `OT_${employee}` : "OT_All") + `_${mm}-${yy}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(fname)}"`);
    res.send(buf);
  } catch (e) {
    res.status(500).send(`Error: ${e.message}`);
  }
});

// ── GET /api/export/payroll/:payId ─────────────────────────
app.get("/api/export/payroll/:payId", requireAdmin, async (req, res) => {
  const { payId } = req.params;
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    const recs   = all.filter(r => r.paidAt === payId);

    if (recs.length === 0) {
      return res.status(404).send(`ไม่พบ records ในรอบ ${payId}`);
    }

    // Get payroll metadata (inline fetch — Payroll_Log!A:J)
    let cutoff = "-", createdAt = "-", createdBy = "Admin";
    try {
      const logRes = await sheets.spreadsheets.values.get({
        spreadsheetId: SHEET_ID,
        range: "Payroll_Log!A3:J5000",
      });
      const logRows = logRes.data.values || [];
      const meta = logRows.find(row => row[0] === payId);
      if (meta) {
        cutoff = meta[1] || "-";
        createdAt = meta[2] || "-";
        createdBy = meta[6] || "Admin";
      }
    } catch (_) { /* tab อาจยังไม่มี */ }

    const wb = XLSX.utils.book_new();

    // Summary
    const byEmp = {};
    recs.forEach(r => {
      byEmp[r.name] = byEmp[r.name] || { name: r.name, days: new Set(), hours: 0, holidays: 0, pay: 0, count: 0 };
      byEmp[r.name].days.add(r.date);
      byEmp[r.name].count += 1;
      byEmp[r.name].pay   += r.pay;
      if (r.otType === "วันธรรมดา") byEmp[r.name].hours += r.hours;
      else                          byEmp[r.name].holidays += 1;
    });
    // ★ v1.25: เพิ่ม travelAllowance + หัก socialSecurity ใน summary
    // ★ v1.30 BUG-05: dedup — ถ้ารับ extras ไปแล้วใน "รอบอื่น" ของเดือนเดียวกัน → travel=0, social=0
    //   (รอบนี้ = payId ปัจจุบัน ต้องนับเสมอ ถึงจะ match กับยอดที่ commit ไป)
    const empListExp = await getEmployees(sheets);
    const empMapExp = {};
    empListExp.forEach(e => empMapExp[e.name] = e);
    const empsPaidExtrasOtherRound = new Set();
    if (cutoff !== "-") {
      try {
        const pl = await sheets.spreadsheets.values.get({
          spreadsheetId: SHEET_ID, range: "Payroll_Log!A3:J5000",
        });
        const cParts = cutoff.split("/");
        const monthKey = `/${cParts[1]}/${cParts[2]}`;
        const otherActivePayIds = new Set();
        const dedupStatuses = new Set(["pending_approval", "approved", "active"]);
        (pl.data.values || []).forEach(row => {
          const pid = (row[0]||"").trim();
          const cs  = (row[1]||"").trim();
          const st  = (row[7]||"active").trim();
          if (pid && pid !== payId && dedupStatuses.has(st) && cs.endsWith(monthKey)) {
            otherActivePayIds.add(pid);
          }
        });
        all.forEach(r => {
          if (r.paidAt && otherActivePayIds.has(r.paidAt)) {
            empsPaidExtrasOtherRound.add(r.name);
          }
        });
      } catch (_) { /* ignore — fallback คือไม่ skip */ }
    }
    const summary = Object.values(byEmp).map(e => {
      const ed = empMapExp[e.name] || {};
      const skipExtras = empsPaidExtrasOtherRound.has(e.name);
      const travel = skipExtras ? 0 : (ed.travelAllowance || 0);
      const social = skipExtras ? 0 : (ed.socialSecurity || 0);
      const salary = skipExtras ? 0 : (ed.salary || 0);    // ★ v1.33
      const netPay = e.pay + travel + salary - social;
      return { name: e.name, days: e.days.size, hours: +e.hours.toFixed(2),
               holidays: e.holidays, pay: e.pay, travel, social, salary, netPay };
    }).sort((a,b) => b.netPay - a.netPay);

    const sumRows = [
      ["ใบรายการจ่าย OT + เงินเดือน"],
      ["รอบจ่าย:", cutoff, "", "Payroll ID:", payId],
      ["จัดทำเมื่อ:", createdAt, "", "ทำโดย:", createdBy],
      [],
      ["ลำดับ","ชื่อพนักงาน","จำนวนวัน","ชั่วโมง","วันหยุด/ตจว.","ค่า OT (฿)","+ ค่าเดินทาง","+ เงินเดือน","- ประกันสังคม","สุทธิ (฿)"],
    ];
    summary.forEach((e, i) => sumRows.push([
      i+1, e.name, e.days, e.hours, e.holidays, e.pay, e.travel, e.salary, e.social, e.netPay,
    ]));
    sumRows.push([]);
    sumRows.push([
      "รวมทั้งหมด", "",
      summary.reduce((s,e)=>s+e.days,0),
      +summary.reduce((s,e)=>s+e.hours,0).toFixed(2),
      summary.reduce((s,e)=>s+e.holidays,0),
      summary.reduce((s,e)=>s+e.pay,0),
      summary.reduce((s,e)=>s+e.travel,0),
      summary.reduce((s,e)=>s+(e.salary||0),0),  // ★ v1.33
      summary.reduce((s,e)=>s+e.social,0),
      summary.reduce((s,e)=>s+e.netPay,0),
    ]);
    const ws1 = XLSX.utils.aoa_to_sheet(sumRows);
    ws1["!cols"] = [{wch:8},{wch:18},{wch:10},{wch:10},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14},{wch:14}];
    ws1["!merges"] = [
      { s: { r: 0, c: 0 }, e: { r: 0, c: 9 } },  // title (10 cols)
    ];
    styleSheet(ws1, { titleRow: 0, headerRow: 4 });
    applyA4Print(ws1, { orientation: "landscape" });
    XLSX.utils.book_append_sheet(wb, ws1, "สรุป");
    setPrintTitles(wb, "สรุป", 4);

    // ── ★ v1.26: Per-employee sheets (1 sheet per คน แบบตัวอย่าง) ──
    const thaiDays = ["อา","จ","อ","พ","พฤ","ศ","ส"];
    function thaiDayOfDate(thaiDateStr) {
      // thaiDateStr = "dd/mm/yyyy" (พ.ศ.)
      try {
        const [d, m, y] = thaiDateStr.split("/").map(Number);
        const dt = new Date(y - 543, m - 1, d);
        return thaiDays[dt.getDay()] || "";
      } catch { return ""; }
    }
    function safeSheetName(name) {
      // Excel: ห้าม : \ / ? * [ ] และยาวไม่เกิน 31 ตัว
      let s = name.replace(/[:\\/?*\[\]]/g, "_").trim();
      if (s.length > 28) s = s.slice(0, 28);
      return s || "พนักงาน";
    }

    // group recs by name
    const recsByName = {};
    recs.forEach(r => {
      (recsByName[r.name] = recsByName[r.name] || []).push(r);
    });

    // sort employees ตามลำดับใน summary (สุทธิมาก→น้อย)
    const usedSheetNames = new Set(["สรุป"]);
    summary.forEach(emp => {
      const list = (recsByName[emp.name] || [])
        .slice()
        .sort((a,b) => a.date.localeCompare(b.date));
      const ed = empMapExp[emp.name] || {};
      const hourlyRate = ed.hourlyRate || 80;

      // นับคืน ตจว. (ค้างคืน) เพื่อใส่ "ตจว. คืนที่ X"
      let opNight = 0;

      const rows = [
        [`ใบรายการ OT — ${emp.name}`],
        [`รอบจ่าย: ${cutoff}    Payroll ID: ${payId}`],
        [],
        ["วันที่","วัน","รายละเอียดงาน","สถานที่","เริ่ม","สิ้นสุด","ชม.","ประเภท","ค่า (฿)"],
      ];

      list.forEach(r => {
        const isOutProv = (r.otType || "").includes("ต่างจังหวัด");
        let startCell = r.startTime;
        let endCell = r.endTime;
        if (isOutProv) {
          opNight += 1;
          startCell = `ตจว. คืนที่ ${opNight}`;
          endCell = "";
        }
        rows.push([
          r.date,
          thaiDayOfDate(r.date),
          r.task || "",
          r.location || "",
          startCell,
          endCell,
          r.hours || 0,
          r.otType,
          r.pay,
        ]);
      });

      // total OT
      const totalOT = list.reduce((s, r) => s + (r.pay || 0), 0);
      const totalHours = list.reduce((s, r) => s + (r.hours || 0), 0);

      rows.push([]);
      rows.push(["", "", "", "", "", "รวม", +totalHours.toFixed(2), "", totalOT]);
      rows.push([]);

      // สรุปจ่าย
      rows.push(["─── สรุปจ่าย ───"]);
      rows.push(["+ ค่า OT", "", "", "", "", "", "", "", emp.pay]);
      if (emp.travel) rows.push(["+ ค่าเดินทาง/พิเศษ", "", "", "", "", "", "", "", emp.travel]);
      if (emp.salary) rows.push(["+ เงินเดือน", "", "", "", "", "", "", "", emp.salary]);   // ★ v1.33
      if (emp.social) rows.push(["- ประกันสังคม", "", "", "", "", "", "", "", -emp.social]);
      rows.push(["ทำจ่ายสุทธิ", "", "", "", "", "", "", "", emp.netPay]);
      rows.push([]);
      rows.push([`Rate: ${hourlyRate} บาท/ชม.    วันหยุด/ตจว.: ${ed.holidayFlat || 0} บาท    ตจว.: ${ed.outProvinceFlat || 0} บาท`]);

      const wsEmp = XLSX.utils.aoa_to_sheet(rows);
      wsEmp["!cols"] = [
        {wch:12},{wch:6},{wch:30},{wch:16},{wch:14},{wch:10},{wch:8},{wch:20},{wch:14},
      ];
      // merge title row + meta row
      wsEmp["!merges"] = [
        { s: { r: 0, c: 0 }, e: { r: 0, c: 8 } },  // ใบรายการ OT — ชื่อ
        { s: { r: 1, c: 0 }, e: { r: 1, c: 8 } },  // รอบจ่าย
      ];
      styleSheet(wsEmp, { titleRow: 0, headerRow: 3 });
      applyA4Print(wsEmp, { orientation: "landscape" });

      // unique sheet name
      let sname = safeSheetName(emp.name);
      let n = 2;
      while (usedSheetNames.has(sname)) {
        sname = safeSheetName(emp.name) + "_" + n;
        n += 1;
      }
      usedSheetNames.add(sname);
      XLSX.utils.book_append_sheet(wb, wsEmp, sname);
      setPrintTitles(wb, sname, 3);
    });

    // ── Details (ทั้งหมดรวมกัน — เก็บไว้สำหรับอ้างอิง) ──
    const detailRows = [
      ["ชื่อ","วันที่","เริ่ม","สิ้นสุด","ชม.","งาน","สถานที่","ประเภท","ค่า (฿)","Payroll ID"],
    ];
    recs.slice().sort((a,b) => a.name.localeCompare(b.name) || a.date.localeCompare(b.date)).forEach(r => {
      detailRows.push([
        r.name, r.date, r.startTime, r.endTime, r.hours,
        r.task||"", r.location||"", r.otType, r.pay, payId,
      ]);
    });
    const ws2 = XLSX.utils.aoa_to_sheet(detailRows);
    ws2["!cols"] = [{wch:18},{wch:13},{wch:8},{wch:8},{wch:8},{wch:28},{wch:16},{wch:18},{wch:12},{wch:20}];
    styleSheet(ws2, { headerRow: 0 });
    applyA4Print(ws2, { orientation: "landscape" });
    XLSX.utils.book_append_sheet(wb, ws2, "รายละเอียดทั้งหมด");
    setPrintTitles(wb, "รายละเอียดทั้งหมด", 0);

    const buf = XLSX.write(wb, { bookType: "xlsx", type: "buffer" });
    const fname = `Payroll_${payId}.xlsx`;
    res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
    res.setHeader("Content-Disposition", `attachment; filename="${encodeURIComponent(fname)}"`);
    res.send(buf);
  } catch (e) {
    res.status(500).send(`Error: ${e.message}`);
  }
});

app.get("/", (_, res) => res.send("🟢 OT Adrun Bot + LIFF running (v1.29)"));

// ── /liff redirect — ถ้ามีคน bookmark URL เก่าไว้ ────────────
app.get("/liff", (_, res) => {
  if (process.env.LIFF_ID) {
    res.redirect(`https://liff.line.me/${process.env.LIFF_ID}`);
  } else {
    res.status(500).send("LIFF_ID env var not configured");
  }
});

// ══════════════════════════════════════════════════════════════
// GOOGLE SHEETS HELPERS  (★ all ranges start at row 3)
// ══════════════════════════════════════════════════════════════
async function getEmployees(sheets) {
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Employees!A3:J500",  // ★ v1.33: เพิ่ม column J (salary)
  });
  return (r.data.values || []).map((row, idx) => ({
    idx,
    name:           (row[0] || "").trim(),
    hourlyRate:     Number(row[1]) || 80,
    holidayFlat:    Number(row[2]) || 500,
    userId:         (row[3] || "").trim(),
    outProvinceFlat: Number(row[4]) || 0,    // ★ v1.25
    travelAllowance: Number(row[5]) || 0,    // ★ v1.25 (รายเดือน)
    socialSecurity:  Number(row[6]) || 0,    // ★ v1.25 (รายเดือน หัก)
    satHalfDay:      String(row[7] || "").toUpperCase() === "TRUE" || row[7] === true || row[7] === "1",
    role:           (row[8] || "").trim().toLowerCase(),  // ★ v1.32
    salary:          Number(row[9]) || 0,    // ★ v1.33: เงินเดือนรายเดือน (จ่ายครั้งเดียวต่อเดือน)
  })).filter(e => e.name);
}

async function getHolidayList(sheets) {
  try {
    const r = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Holidays!A3:C500",  // ★ v1.14: ขยายจาก 100 → 500 (~25 ปี)
    });
    return (r.data.values || []).map((row, idx) => ({
      idx,
      date: (row[0]||"").trim(),
      day:  (row[1]||"").trim(),
      name: (row[2]||"").trim(),
    })).filter(h => h.date);
  } catch (_) { return []; }
}

async function getAllRecords(sheets) {
  const r = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "OT_Records!A3:K20000",  // ★ v1.15: เพิ่ม column K (paidAt)
  });
  return (r.data.values || []).map((row, i) => ({
    idx: i, name: row[0]||"", date: row[1]||"",
    startTime: row[2]||"-", endTime: row[3]||"-",
    hours: Number(row[4])||0, task: row[5]||"", location: row[6]||"",
    otType: row[7]||"", pay: Number(row[8])||0, createdAt: row[9]||"",
    paidAt: (row[10]||"").trim(),  // ★ v1.15: ID รอบจ่ายเงิน (เช่น "PAY-20260528-1430")
  }));
}

async function getDayHours(sheets, name, date) {
  const all = await getAllRecords(sheets);
  return all.filter(r => r.name === name && r.date === date && r.otType === "วันธรรมดา")
            .reduce((s, r) => s + r.hours, 0);
}

// ★ v1.21+v1.33: ตรวจช่วงเวลาทับกับ record อื่น
//   v1.33: ครอบคลุม cross-day overlap — record วันก่อนหน้าที่ข้ามวันมาทับ + record วันถัดไปที่ shift ไปต้น
async function findOverlappingRecord(sheets, name, date, newStart, newEnd) {
  const all = await getAllRecords(sheets);
  const toMin = t => { const [h, m] = t.split(":").map(Number); return h * 60 + m; };

  // หาวันก่อน + วันถัดไป (เพื่อครอบคลุม cross-day record)
  const [dd, mm, yy] = date.split("/").map(Number);
  const dt = new Date(yy - 543, mm - 1, dd);
  const fmt = d => `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()+543}`;
  const prevDate = fmt(new Date(dt.getTime() - 86400000));
  const nextDate = fmt(new Date(dt.getTime() + 86400000));

  const candidates = all.filter(r =>
    r.name === name &&
    (r.date === date || r.date === prevDate || r.date === nextDate) &&
    r.otType === "วันธรรมดา" &&
    r.startTime !== "-" && r.endTime !== "-"
  );

  // แปลง new range → absolute timeline (อิงวันของ record ใหม่)
  let nS = toMin(newStart);
  let nE = toMin(newEnd);
  if (nE < nS) nE += 1440; // cross-midnight

  for (const r of candidates) {
    let s = toMin(r.startTime);
    let e = toMin(r.endTime);
    if (e < s) e += 1440; // record เก่าข้ามคืน
    // Shift r ตามวัน — ให้อยู่บน timeline เดียวกับ new
    if (r.date === prevDate)      { s -= 1440; e -= 1440; }
    else if (r.date === nextDate) { s += 1440; e += 1440; }
    // เช็ค overlap ปกติ
    if (nS < e && s < nE) return r;
  }
  return null;
}

async function saveRecord(sheets, data) {
  const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "OT_Records!A:J",
    valueInputOption: "USER_ENTERED",
    resource: { values: [[
      data.name, data.date, data.startTime, data.endTime,
      data.hours, data.task, data.location || "",
      data.otType, data.pay, now,
    ]] },
  });
}

// ★ ลบ row จริงด้วย batchUpdate (ไม่ใช่ clear)
async function deleteRow(sheets, sheetName, rowIndex1Based) {
  const meta  = await sheets.spreadsheets.get({ spreadsheetId: SHEET_ID });
  const sheet = meta.data.sheets.find(s => s.properties.title === sheetName);
  if (!sheet) throw new Error(`ไม่พบชีต "${sheetName}"`);
  const sheetId = sheet.properties.sheetId;

  await sheets.spreadsheets.batchUpdate({
    spreadsheetId: SHEET_ID,
    resource: {
      requests: [{
        deleteDimension: {
          range: {
            sheetId,
            dimension:  "ROWS",
            startIndex: rowIndex1Based - 1,  // 0-indexed inclusive
            endIndex:   rowIndex1Based,      // exclusive
          },
        },
      }],
    },
  });
}

function calcHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  let mins = eh * 60 + em - sh * 60 - sm;
  // ★ v1.4: ข้ามวัน (เช่น 22:00–04:00) → +24 ชม.
  if (mins < 0) mins += 24 * 60;
  return mins > 0 ? +(mins / 60).toFixed(2) : 0;
}

// ★ v1.3+v1.4: เช็คว่า OT ช่วงเวลานี้ทับเวลางานปกติ (08:30–17:30) — รองรับข้ามวัน
// ★ v1.27: รับ workEndMin ปรับได้ (เช่น เสาร์ครึ่งวัน → 12:00)
function overlapsWorkHours(startTime, endTime, workEndMin = WORK_END_MIN) {
  const [sh, sm] = startTime.split(":").map(Number);
  const [eh, em] = endTime.split(":").map(Number);
  let s = sh * 60 + sm;
  let e = eh * 60 + em;
  if (e < s) e += 24 * 60; // ข้ามวัน
  // ตรวจทับเวลางานทั้งวันแรก และวันถัดไป (กรณี OT ข้ามวัน)
  const day1 = s < workEndMin          && e > WORK_START_MIN;
  const day2 = s < workEndMin + 1440   && e > WORK_START_MIN + 1440;
  return day1 || day2;
}

// ★ v1.27: คำนวณ workEndMin ตามพนักงาน + วัน (เสาร์ครึ่งวัน → 12:00)
function getWorkEndForEmp(emp, dow) {
  if (emp && emp.satHalfDay && dow === 6) return WORK_END_MIN_SAT_HALF;
  return WORK_END_MIN;
}
// label สำหรับแสดง error
function workWindowLabel(workEndMin) {
  const eh = Math.floor(workEndMin / 60);
  const em = workEndMin % 60;
  return `08:30–${String(eh).padStart(2,"0")}:${String(em).padStart(2,"0")}`;
}

// ★ v1.30: ลบ validateOTWindow (dead code — ไม่เคยถูกเรียก)
//   overlapsWorkHours คุมขอบเขตอยู่แล้ว ถ้าจะคืน rule 06:00-06:00 ค่อย wire กลับ

function getTodayThai() {
  const d = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Bangkok" }));
  return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()+543}`;
}

// ★ NEW: คำนวณ "วันในสัปดาห์" จาก dd/mm/yyyy(พ.ศ.)
function getDowFullThai(thaiDateStr) {
  const [dd, mm, yy] = String(thaiDateStr).split("/");
  const d = new Date(Number(yy) - 543, Number(mm) - 1, Number(dd));
  return ["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"][d.getDay()] || "";
}

// ══════════════════════════════════════════════════════════════
// LINE BOT EVENT HANDLER
// ══════════════════════════════════════════════════════════════
async function handleBotEvent(event) {
  if (event.type !== "message" || event.message.type !== "text") return;
  const text  = event.message.text.trim();
  const lower = text.toLowerCase();

  // ★ v1.34: คำสั่ง #คิว — ส่งลิงค์ระบบคิวงาน
  if (lower === "#คิว" || lower === "#queue") {
    return client.replyMessage(event.replyToken, {
      type: "template",
      altText: "📋 ระบบคิวงาน Adrun",
      template: {
        type: "buttons",
        title: "📋 คิวงาน Adrun",
        text: "กดปุ่มด้านล่างเพื่อดูคิวงาน",
        actions: [{ type: "uri", label: "🔗 เปิดคิวงาน", uri: "https://work.adrun.co.th" }],
      },
    });
  }

  if (!lower.startsWith("#ot") && !lower.startsWith("#โอที")) return;

  // ★ ใช้ canonical LIFF URL — LINE จะ redirect ไป Endpoint URL ที่ตั้งไว้เอง
  const liffUrl = process.env.LIFF_ID
    ? `https://liff.line.me/${process.env.LIFF_ID}`
    : `https://${process.env.RAILWAY_PUBLIC_DOMAIN || "your-app.up.railway.app"}`;

  if (lower === "#ot" || lower.includes("เปิด") || lower.includes("บันทึก")) {
    return client.replyMessage(event.replyToken, {
      type: "template",
      altText: "เปิดระบบบันทึก OT",
      template: {
        type: "buttons",
        title: "🟢 ระบบ OT บริษัท Adrun",
        text: "กดปุ่มด้านล่างเพื่อเปิดฟอร์มบันทึก OT",
        actions: [{ type: "uri", label: "📋 เปิดระบบ OT", uri: liffUrl }],
      },
    });
  }

  try {
    const groupId = event.source.groupId;
    const userId  = event.source.userId;
    let senderName = "ไม่ทราบชื่อ";
    try {
      const profile = groupId
        ? await client.getGroupMemberProfile(groupId, userId)
        : await client.getProfile(userId);
      senderName = profile.displayName;
    } catch (_) {}

    const sheets    = await getSheetsClient();
    const employees = await getEmployees(sheets);

    // ★ B+ patch: หาด้วย userId ก่อน → fallback displayName
    let empData = employees.find(e => e.userId && e.userId === userId);
    if (!empData) empData = employees.find(e => e.name === senderName);

    if (lower.includes("สรุป")) {
      return client.replyMessage(event.replyToken, await buildSummary(sheets, empData?.name || senderName));
    }
    if (lower.includes("ช่วย") || lower.includes("help")) {
      return client.replyMessage(event.replyToken, {
        type: "text",
        text: `📖 วิธีใช้ระบบ OT\n\nกด #OT เพื่อเปิดฟอร์มบันทึก OT\nหรือพิมพ์:\n#OT 18:00 21:00 งาน | สถานที่\n#OT วันหยุด งาน\n#OT สรุป`,
      });
    }

    if (!empData) {
      return client.replyMessage(event.replyToken, { type:"text", text:`⚠️ ไม่พบชื่อ "${senderName}" กรุณาแจ้ง Admin` });
    }

    const holidays  = await getHolidayList(sheets);
    const todayDate = getTodayThai();
    const todayDow  = new Date().getDay();
    const isHolDate = holidays.some(h => h.date === todayDate);
    const isHolCmd  = lower.includes("วันหยุด") || lower.includes("หยุด");

    if (isHolCmd || isHolDate || todayDow === 0) {
      const parts    = text.replace(/#OT/i,"").replace(/วันหยุด|หยุด/g,"").trim();
      const [task="", location=""] = parts.includes("|") ? parts.split("|").map(s=>s.trim()) : [parts, ""];
      const typeLabel = todayDow===0 ? "วันอาทิตย์" : isHolDate ? "วันหยุดนักขัตฤกษ์" : "วันหยุด";
      // ★ v1.30 BUG-04: wrap save ด้วย mutex
      await withOTMutex(`${empData.name}|${todayDate}`, () =>
        saveRecord(sheets, { name:empData.name, date:todayDate, startTime:"-", endTime:"-", hours:0, task:task||"-", location, otType:typeLabel, pay:empData.holidayFlat })
      );
      return client.replyMessage(event.replyToken, { type:"text", text:`✅ บันทึก OT ${typeLabel}\n👤 ${empData.name}\n📝 ${task||"-"}\n📅 ${todayDate}` });
    }

    const times = [...text.matchAll(/\b(\d{1,2}):(\d{2})\b/g)];
    if (times.length < 2) return client.replyMessage(event.replyToken, { type:"text", text:`❓ รูปแบบผิด\nลอง: #OT 18:00 21:00 งาน\nหรือกด #OT เพื่อเปิดฟอร์ม` });

    const [startTime, endTime] = [times[0][0], times[1][0]];
    const hours    = calcHours(startTime, endTime);
    if (hours <= 0) return client.replyMessage(event.replyToken, { type:"text", text:"⚠️ เวลาไม่ถูกต้อง" });

    // ★ v1.6: ลงเวลาได้ทุกช่วง ห้ามแค่ทับเวลางาน
    // ★ v1.27: เสาร์ครึ่งวัน → window 08:30-12:00
    const botWorkEnd = getWorkEndForEmp(empData, todayDow);
    if (overlapsWorkHours(startTime, endTime, botWorkEnd)) {
      return client.replyMessage(event.replyToken, { type:"text", text:`⚠️ ช่วง ${workWindowLabel(botWorkEnd)} เป็นเวลางานปกติ ไม่สามารถบันทึก OT ได้` });
    }

    // ★ v1.30 BUG-04: ทุก check + save ทำใน mutex (atomic ต่อ name|date)
    const result = await withOTMutex(`${empData.name}|${todayDate}`, async () => {
      // ★ v1.21: ห้ามทับกับ record ในวันเดียวกัน
      const overlap = await findOverlappingRecord(sheets, empData.name, todayDate, startTime, endTime);
      if (overlap) return { ok: false, reason: `⚠️ ช่วง ${startTime}-${endTime} ทับกับรายการเดิม ${overlap.startTime}-${overlap.endTime} (${overlap.hours} ชม.) ในวันเดียวกัน` };
      // ★ v1.2: ลงเวลาตามจริง คำนวณค่า OT สูงสุด MAX_OT_PER_DAY ชม./วัน
      const already = await getDayHours(sheets, empData.name, todayDate);
      const remainingPayable = Math.max(0, MAX_OT_PER_DAY - already);
      const payableHours     = Math.min(hours, remainingPayable);
      const after = text.replace(/#OT/i,"").replace(startTime,"").replace(endTime,"").trim();
      const [task="", location=""] = after.includes("|") ? after.split("|").map(s=>s.trim()) : [after, ""];
      const pay   = Math.round(payableHours * empData.hourlyRate * WEEKDAY_MULTIPLIER);
      await saveRecord(sheets, { name:empData.name, date:todayDate, startTime, endTime, hours, task:task||"-", location, otType:"วันธรรมดา", pay });
      return { ok: true, payableHours, task };
    });

    if (!result.ok) {
      return client.replyMessage(event.replyToken, { type:"text", text: result.reason });
    }
    const replyText = result.payableHours < hours
      ? `✅ บันทึก OT\n👤 ${empData.name}\n⏰ ${startTime}–${endTime}\n📊 ทำจริง ${hours}ชม. · คิด ${result.payableHours}ชม.\n📝 ${result.task||"-"}\n📅 ${todayDate}`
      : `✅ บันทึก OT\n👤 ${empData.name}\n⏰ ${startTime}–${endTime} (${hours}ชม.)\n📝 ${result.task||"-"}\n📅 ${todayDate}`;
    return client.replyMessage(event.replyToken, { type:"text", text: replyText });

  } catch (err) {
    console.error(err);
    return client.replyMessage(event.replyToken, { type:"text", text:"❌ เกิดข้อผิดพลาด" });
  }
}

async function buildSummary(sheets, name) {
  const all  = await getAllRecords(sheets);
  const d    = new Date(new Date().toLocaleString("en-US",{timeZone:"Asia/Bangkok"}));
  const mm   = String(d.getMonth()+1).padStart(2,"0");
  const yy   = String(d.getFullYear()+543);
  const mine = all.filter(r => r.name===name && r.date.endsWith(`/${mm}/${yy}`));
  if (!mine.length) return { type:"text", text:`📊 ${name}\nยังไม่มี OT เดือน ${mm}/${yy}` };
  const h  = mine.filter(r=>r.otType==="วันธรรมดา").reduce((s,r)=>s+r.hours,0);
  const hl = mine.filter(r=>r.otType!=="วันธรรมดา").length;
  return { type:"text", text:`📊 ${name} เดือน ${mm}/${yy}\n⏱ ${h} ชม.\n🌅 วันหยุด ${hl} วัน\nรายการ ${mine.length} รายการ` };
}

// ── Start ─────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🟢 OT Bot + LIFF on port ${PORT} (v1.1 B+)`));
