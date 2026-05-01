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

// ── LINE Config ──────────────────────────────────────────────
const lineConfig = {
  channelAccessToken: process.env.LINE_TOKEN,
  channelSecret:      process.env.LINE_SECRET,
};
const client = new line.Client(lineConfig);
const app    = express();

// ── Webhook ต้องมาก่อน express.json() เสมอ ──────────────────
app.post("/webhook", express.raw({ type: "*/*" }), async (req, res) => {
  console.log("📨 Webhook received! sig=", req.headers["x-line-signature"]?.slice(0, 20));
  res.json({ status: "ok" });
  const body = req.body;
  const sig  = req.headers["x-line-signature"];
  const hash = crypto.createHmac("SHA256", process.env.LINE_SECRET)
                     .update(body).digest("base64");
  if (sig !== hash) {
    console.error("❌ Signature mismatch! LINE_SECRET ผิด หรือไม่ตรงกับ Messaging API channel");
    console.error("   expected:", hash.slice(0, 20));
    console.error("   received:", sig?.slice(0, 20));
    return;
  }
  console.log("✅ Signature OK");
  const events = JSON.parse(body.toString()).events || [];
  console.log(`📋 Events: ${events.length}`, events.map(e => `${e.type}:${e.message?.text || ""}`));
  await Promise.all(events.map(handleBotEvent));
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

// ── OT Rules ─────────────────────────────────────────────────
const MAX_OT_PER_DAY     = 5;
const WEEKDAY_MULTIPLIER = 1.5;

// ★ v1.3: เวลางานปกติ จันทร์–เสาร์ (ห้ามลง OT ทับ)
const WORK_START_MIN = 8 * 60 + 30;   // 08:30 = 510
const WORK_END_MIN   = 17 * 60 + 30;  // 17:30 = 1050

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

// ── GET /api/me?userId=Uxxxx&displayName=xxx ────────────────
// ★ B+ patch: match by userId first → fall back displayName + auto-bind
app.get("/api/me", async (req, res) => {
  const { userId, displayName } = req.query;
  try {
    const sheets    = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const admins    = (process.env.ADMIN_LINE_IDS || "")
                        .split(",").map(s => s.trim()).filter(Boolean);

    // 1) จับคู่ด้วย userId ก่อน (น่าเชื่อถือสุด)
    let emp = employees.find(e => e.userId && e.userId === userId);
    let matchedBy = emp ? "userId" : null;

    // 2) ถ้าไม่เจอ → fallback หาด้วย displayName
    if (!emp && displayName) {
      emp = employees.find(e => e.name === displayName);
      if (emp) {
        matchedBy = "displayName";
        // 3) auto-bind: เจอด้วยชื่อ + ยังไม่มี userId → เขียนกลับ
        if (!emp.userId && userId) {
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
        }
      }
    }

    const isAdmin = admins.includes(userId);

    res.json({
      found:       !!emp,
      name:        emp?.name || displayName || "",
      userId,
      isAdmin,
      hourlyRate:  emp?.hourlyRate  || 0,
      holidayFlat: emp?.holidayFlat || 0,
      matchedBy,
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
app.post("/api/employees", async (req, res) => {
  const { name, hourlyRate, holidayFlat, userId } = req.body;
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "Employees!A:D",
      valueInputOption: "USER_ENTERED",
      resource: { values: [[name, hourlyRate, holidayFlat, userId || ""]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── PUT /api/employees/:idx — แก้ไขพนักงาน (Admin) ──────────
app.put("/api/employees/:idx", async (req, res) => {
  const row = idxToRow(Number(req.params.idx));
  const { name, hourlyRate, holidayFlat, userId } = req.body;
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Employees!A${row}:D${row}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[name, hourlyRate, holidayFlat, userId || ""]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── DELETE /api/employees/:idx — ลบพนักงาน (Admin) ★ NEW ────
app.delete("/api/employees/:idx", async (req, res) => {
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
app.post("/api/holidays", async (req, res) => {
  const { date, name: hName } = req.body;
  try {
    const sheets = await getSheetsClient();
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
app.put("/api/holidays/:idx", async (req, res) => {
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
app.delete("/api/holidays/:idx", async (req, res) => {
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
app.post("/api/ot", async (req, res) => {
  const { name, date, startTime, endTime, task, location, otType } = req.body;
  try {
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

    // ★ v1.6: ลงเวลาได้ทุกช่วง (รวม 01:00-04:00 = OT คืนวันนั้น) ห้ามแค่ทับเวลางาน 08:30–17:30
    if (overlapsWorkHours(startTime, endTime)) {
      return res.status(400).json({ error: "ช่วง 08:30–17:30 เป็นเวลางานปกติ ไม่สามารถบันทึก OT ได้" });
    }

    // ★ v1.2: บันทึกเวลาตามจริง แต่คำนวณค่า OT สูงสุด MAX_OT_PER_DAY ชม./วัน
    const alreadyDay        = await getDayHours(sheets, name, date);
    const remainingPayable  = Math.max(0, MAX_OT_PER_DAY - alreadyDay);
    const payableHours      = Math.min(hours, remainingPayable);
    const pay = Math.round(payableHours * emp.hourlyRate * WEEKDAY_MULTIPLIER);

    await saveRecord(sheets, { name, date, startTime, endTime, hours, task, location, otType: "วันธรรมดา", pay });
    return res.json({ ok: true, hours, payableHours, pay, otType: "วันธรรมดา", capped: payableHours < hours });

  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/bind-employee — Self-claim ผูกบัญชี LINE ★ v1.13
app.post("/api/bind-employee", async (req, res) => {
  const { userId, employeeName } = req.body;
  if (!userId || !employeeName) return res.status(400).json({ error: "ข้อมูลไม่ครบ" });
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
app.post("/api/edit-request", async (req, res) => {
  const { name, date, recordDesc, note } = req.body;
  try {
    const sheets = await getSheetsClient();
    const now    = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
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

// ── GET /api/admin/records — ทุก record สำหรับ Admin ─────────
app.get("/api/admin/records", async (req, res) => {
  const { month, year } = req.query;
  try {
    const sheets = await getSheetsClient();
    const all    = await getAllRecords(sheets);
    const rows   = all.filter(r =>
      !month || r.date.includes(`/${month}/${year}`)
    );
    res.json({ records: rows });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/edit-requests — คำขอแก้ไข (Admin) ──────────────
app.get("/api/edit-requests", async (req, res) => {
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
app.put("/api/edit-requests/:idx", async (req, res) => {
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

app.get("/", (_, res) => res.send("🟢 OT Adrun Bot + LIFF running (v1.1)"));

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
    range: "Employees!A3:D500",  // ★ v1.14: ขยายจาก 200 → 500 (รองรับ 500 คน)
  });
  return (r.data.values || []).map((row, idx) => ({
    idx,
    name:        (row[0] || "").trim(),
    hourlyRate:  Number(row[1]) || 80,
    holidayFlat: Number(row[2]) || 500,
    userId:      (row[3] || "").trim(),
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
    range: "OT_Records!A3:J20000",  // ★ v1.14: ขยายจาก 2000 → 20000 (~12 ปี)
  });
  return (r.data.values || []).map((row, i) => ({
    idx: i, name: row[0]||"", date: row[1]||"",
    startTime: row[2]||"-", endTime: row[3]||"-",
    hours: Number(row[4])||0, task: row[5]||"", location: row[6]||"",
    otType: row[7]||"", pay: Number(row[8])||0, createdAt: row[9]||"",
  }));
}

async function getDayHours(sheets, name, date) {
  const all = await getAllRecords(sheets);
  return all.filter(r => r.name === name && r.date === date && r.otType === "วันธรรมดา")
            .reduce((s, r) => s + r.hours, 0);
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
function overlapsWorkHours(startTime, endTime) {
  const [sh, sm] = startTime.split(":").map(Number);
  const [eh, em] = endTime.split(":").map(Number);
  let s = sh * 60 + sm;
  let e = eh * 60 + em;
  if (e < s) e += 24 * 60; // ข้ามวัน
  // ตรวจทับเวลางานทั้งวันแรก และวันถัดไป (กรณี OT ข้ามวัน)
  const day1 = s < WORK_END_MIN          && e > WORK_START_MIN;
  const day2 = s < WORK_END_MIN + 1440   && e > WORK_START_MIN + 1440;
  return day1 || day2;
}

// ★ v1.5: เช็คว่า OT อยู่ในช่วง OT day (06:00 ถึง 06:00 ถัดไป) หรือเปล่า
function validateOTWindow(startTime, endTime) {
  const [sh, sm] = startTime.split(":").map(Number);
  const [eh, em] = endTime.split(":").map(Number);
  let s = sh * 60 + sm;
  let e = eh * 60 + em;
  if (e < s) e += 24 * 60; // ข้ามวัน
  if (s < OT_DAY_START_MIN) return "เวลาเริ่ม OT ต้องไม่ก่อน 06:00";
  if (e > OT_DAY_END_MIN)   return "OT ต้องสิ้นสุดไม่เกิน 06:00 ของวันถัดไป";
  return null;
}

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
      await saveRecord(sheets, { name:empData.name, date:todayDate, startTime:"-", endTime:"-", hours:0, task:task||"-", location, otType:typeLabel, pay:empData.holidayFlat });
      return client.replyMessage(event.replyToken, { type:"text", text:`✅ บันทึก OT ${typeLabel}\n👤 ${empData.name}\n📝 ${task||"-"}\n📅 ${todayDate}` });
    }

    const times = [...text.matchAll(/\b(\d{1,2}):(\d{2})\b/g)];
    if (times.length < 2) return client.replyMessage(event.replyToken, { type:"text", text:`❓ รูปแบบผิด\nลอง: #OT 18:00 21:00 งาน\nหรือกด #OT เพื่อเปิดฟอร์ม` });

    const [startTime, endTime] = [times[0][0], times[1][0]];
    const hours    = calcHours(startTime, endTime);
    const already  = await getDayHours(sheets, empData.name, todayDate);
    if (hours <= 0) return client.replyMessage(event.replyToken, { type:"text", text:"⚠️ เวลาไม่ถูกต้อง" });

    // ★ v1.6: ลงเวลาได้ทุกช่วง ห้ามแค่ทับเวลางาน 08:30–17:30
    if (overlapsWorkHours(startTime, endTime)) {
      return client.replyMessage(event.replyToken, { type:"text", text:"⚠️ ช่วง 08:30–17:30 เป็นเวลางานปกติ ไม่สามารถบันทึก OT ได้" });
    }

    // ★ v1.2: ลงเวลาตามจริง คำนวณค่า OT สูงสุด MAX_OT_PER_DAY ชม./วัน
    const remainingPayable = Math.max(0, MAX_OT_PER_DAY - already);
    const payableHours     = Math.min(hours, remainingPayable);

    const after = text.replace(/#OT/i,"").replace(startTime,"").replace(endTime,"").trim();
    const [task="", location=""] = after.includes("|") ? after.split("|").map(s=>s.trim()) : [after, ""];
    const pay   = Math.round(payableHours * empData.hourlyRate * WEEKDAY_MULTIPLIER);
    await saveRecord(sheets, { name:empData.name, date:todayDate, startTime, endTime, hours, task:task||"-", location, otType:"วันธรรมดา", pay });

    const replyText = payableHours < hours
      ? `✅ บันทึก OT\n👤 ${empData.name}\n⏰ ${startTime}–${endTime}\n📊 ทำจริง ${hours}ชม. · คิดเงิน ${payableHours}ชม. (${pay}฿)\n📝 ${task||"-"}\n📅 ${todayDate}`
      : `✅ บันทึก OT\n👤 ${empData.name}\n⏰ ${startTime}–${endTime} (${hours}ชม. · ${pay}฿)\n📝 ${task||"-"}\n📅 ${todayDate}`;
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
