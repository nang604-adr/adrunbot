// ============================================================
// OT LINE Bot + LIFF Server — index.js
// บริษัท Adrun
// LINE Messaging API (Bot) + LIFF REST API + Static file serve
// ============================================================
require("dotenv").config();
const express    = require("express");
const line       = require("@line/bot-sdk");
const { google } = require("googleapis");
const path       = require("path");
const https      = require("https");

// ── LINE Config ──────────────────────────────────────────────
const lineConfig = {
  channelAccessToken: process.env.LINE_TOKEN,
  channelSecret:      process.env.LINE_SECRET,
};
const client = new line.Client(lineConfig);
const app    = express();
const crypto = require("crypto");

// ── Webhook ต้องมาก่อน express.json() เสมอ ──────────────────
app.post("/webhook", express.raw({ type: "*/*" }), async (req, res) => {
  res.json({ status: "ok" });
  const body   = req.body;
  const sig    = req.headers["x-line-signature"];
  const hash   = crypto.createHmac("SHA256", process.env.LINE_SECRET)
                        .update(body).digest("base64");
  if (sig !== hash) return;
  const events = JSON.parse(body.toString()).events || [];
  await Promise.all(events.map(handleBotEvent));
});

// ── Middleware ───────────────────────────────────────────────
app.use(express.json());
app.use(express.static(path.join(__dirname, "public"))); // serve LIFF HTML

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

// ══════════════════════════════════════════════════════════════
// LIFF REST API ENDPOINTS
// ══════════════════════════════════════════════════════════════

// ── GET /api/config — ส่ง LIFF ID ไปให้ front-end ──────────
app.get("/api/config", (_, res) => {
  res.json({ liffId: process.env.LIFF_ID });
});

// ── GET /api/me?userId=Uxxxx — ดึงข้อมูลพนักงาน ────────────
app.get("/api/me", async (req, res) => {
  const { userId, displayName } = req.query;
  try {
    const sheets    = await getSheetsClient();
    const employees = await getEmployees(sheets);
    const admins    = (process.env.ADMIN_LINE_IDS || "").split(",").map(s => s.trim());

    const emp = employees.find(e => e.name === displayName);
    const isAdmin = admins.includes(userId);

    res.json({
      found:       !!emp,
      name:        displayName,
      isAdmin,
      hourlyRate:  emp?.hourlyRate  || 0,
      holidayFlat: emp?.holidayFlat || 0,
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

// ── GET /api/employees — รายชื่อพนักงานทั้งหมด (Admin) ──────
app.get("/api/employees", async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    res.json({ employees: await getEmployees(sheets) });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── GET /api/holidays — วันหยุดประจำปี ──────────────────────
app.get("/api/holidays", async (req, res) => {
  try {
    const sheets = await getSheetsClient();
    res.json({ holidays: await getHolidayList(sheets) });
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
    const jsDate   = new Date(Number(yy) - 543, Number(mm) - 1, Number(dd));
    const dow      = jsDate.getDay();
    const isHolDate = holidays.some(h => h.date === date);
    const isSun    = dow === 0;
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
    if (hours <= 0)          return res.status(400).json({ error: "เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น" });
    if (hours > MAX_OT_PER_DAY) return res.status(400).json({ error: `OT สูงสุด ${MAX_OT_PER_DAY} ชม./วัน` });

    const alreadyDay = await getDayHours(sheets, name, date);
    if (alreadyDay + hours > MAX_OT_PER_DAY) {
      const remain = +(MAX_OT_PER_DAY - alreadyDay).toFixed(1);
      return res.status(400).json({ error: `วันนี้บันทึกไปแล้ว ${alreadyDay} ชม. เพิ่มได้อีก ${remain} ชม.` });
    }

    const pay = Math.round(hours * emp.hourlyRate * WEEKDAY_MULTIPLIER);
    await saveRecord(sheets, { name, date, startTime, endTime, hours, task, location, otType: "วันธรรมดา", pay });
    return res.json({ ok: true, hours, pay, otType: "วันธรรมดา" });

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

// ── PUT /api/employees/:idx — แก้ไขพนักงาน (Admin) ──────────
app.put("/api/employees/:idx", async (req, res) => {
  const idx = Number(req.params.idx) + 2; // row = idx + 2 (1 header + 1-indexed)
  const { name, hourlyRate, holidayFlat } = req.body;
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.update({
      spreadsheetId: SHEET_ID,
      range: `Employees!A${idx}:C${idx}`,
      valueInputOption: "USER_ENTERED",
      resource: { values: [[name, hourlyRate, holidayFlat]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/employees — เพิ่มพนักงาน (Admin) ──────────────
app.post("/api/employees", async (req, res) => {
  const { name, hourlyRate, holidayFlat } = req.body;
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "Employees!A:C",
      valueInputOption: "USER_ENTERED",
      resource: { values: [[name, hourlyRate, holidayFlat]] },
    });
    res.json({ ok: true });
  } catch (e) {
    res.status(500).json({ error: e.message });
  }
});

// ── POST /api/holidays — เพิ่มวันหยุด (Admin) ───────────────
app.post("/api/holidays", async (req, res) => {
  const { date, name: hName } = req.body;
  try {
    const sheets = await getSheetsClient();
    await sheets.spreadsheets.values.append({
      spreadsheetId: SHEET_ID,
      range: "Holidays!A:C",
      valueInputOption: "USER_ENTERED",
      resource: { values: [[date, "", hName]] },
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
      spreadsheetId: SHEET_ID, range: "Edit_Requests!A2:F500",
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
  const row    = Number(req.params.idx) + 2;
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

app.get("/", (_, res) => res.send("🟢 OT Adrun Bot + LIFF running"));

// ══════════════════════════════════════════════════════════════
// GOOGLE SHEETS HELPERS
// ══════════════════════════════════════════════════════════════
async function getEmployees(sheets) {
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "Employees!A2:C200" });
  return (r.data.values || []).map(row => ({
    name: (row[0] || "").trim(), hourlyRate: Number(row[1]) || 80, holidayFlat: Number(row[2]) || 500,
  })).filter(e => e.name);
}

async function getHolidayList(sheets) {
  try {
    const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "Holidays!A2:C100" });
    return (r.data.values || []).map(row => ({ date: (row[0]||"").trim(), name: (row[2]||row[1]||"").trim() })).filter(h => h.date);
  } catch (_) { return []; }
}

async function getAllRecords(sheets) {
  const r = await sheets.spreadsheets.values.get({ spreadsheetId: SHEET_ID, range: "OT_Records!A2:J2000" });
  return (r.data.values || []).map((row, i) => ({
    idx: i, name: row[0]||"", date: row[1]||"", startTime: row[2]||"-", endTime: row[3]||"-",
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
    spreadsheetId: SHEET_ID, range: "OT_Records!A:J", valueInputOption: "USER_ENTERED",
    resource: { values: [[data.name, data.date, data.startTime, data.endTime,
      data.hours, data.task, data.location||"", data.otType, data.pay, now]] },
  });
}

function calcHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  const mins = eh * 60 + em - sh * 60 - sm;
  return mins > 0 ? +(mins / 60).toFixed(2) : 0;
}

function getTodayThai() {
  const d  = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Bangkok" }));
  return `${String(d.getDate()).padStart(2,"0")}/${String(d.getMonth()+1).padStart(2,"0")}/${d.getFullYear()+543}`;
}

// ── Bot event handler (ย่อ — เหมือน index.js เดิม) ──────────
async function handleBotEvent(event) {
  if (event.type !== "message" || event.message.type !== "text") return;
  const text  = event.message.text.trim();
  const lower = text.toLowerCase();
  if (!lower.startsWith("#ot") && !lower.startsWith("#โอที")) return;

  const liffUrl = `https://${process.env.RAILWAY_PUBLIC_DOMAIN || "your-app.up.railway.app"}/liff`;

  // ถ้าพิมพ์ #OT หรือ #OT เปิด → ส่ง LIFF link
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

  // ยังรองรับการพิมพ์คำสั่งแบบเดิมด้วย
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
    const empData   = employees.find(e => e.name === senderName);

    if (lower.includes("สรุป")) {
      return client.replyMessage(event.replyToken, await buildSummary(sheets, senderName));
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

    const holidays     = await getHolidayList(sheets);
    const todayDate    = getTodayThai();
    const todayDow     = new Date().getDay();
    const isHolDate    = holidays.some(h => h.date === todayDate);
    const isHolCmd     = lower.includes("วันหยุด") || lower.includes("หยุด");

    if (isHolCmd || isHolDate || todayDow === 0) {
      const parts    = text.replace(/#OT/i,"").replace(/วันหยุด|หยุด/g,"").trim();
      const [task="", location=""] = parts.includes("|") ? parts.split("|").map(s=>s.trim()) : [parts, ""];
      const typeLabel = todayDow===0 ? "วันอาทิตย์" : isHolDate ? "วันหยุดนักขัตฤกษ์" : "วันหยุด";
      await saveRecord(sheets, { name:senderName, date:todayDate, startTime:"-", endTime:"-", hours:0, task:task||"-", location, otType:typeLabel, pay:empData.holidayFlat });
      return client.replyMessage(event.replyToken, { type:"text", text:`✅ บันทึก OT ${typeLabel}\n👤 ${senderName}\n📝 ${task||"-"}\n📅 ${todayDate}` });
    }

    const times = [...text.matchAll(/\b(\d{1,2}):(\d{2})\b/g)];
    if (times.length < 2) return client.replyMessage(event.replyToken, { type:"text", text:`❓ รูปแบบผิด\nลอง: #OT 18:00 21:00 งาน\nหรือกด #OT เพื่อเปิดฟอร์ม` });

    const [startTime, endTime] = [times[0][0], times[1][0]];
    const hours    = calcHours(startTime, endTime);
    const already  = await getDayHours(sheets, senderName, todayDate);
    if (hours <= 0)                          return client.replyMessage(event.replyToken, { type:"text", text:"⚠️ เวลาไม่ถูกต้อง" });
    if (hours > MAX_OT_PER_DAY)              return client.replyMessage(event.replyToken, { type:"text", text:`⚠️ เกิน ${MAX_OT_PER_DAY} ชม./วัน` });
    if (already + hours > MAX_OT_PER_DAY)    return client.replyMessage(event.replyToken, { type:"text", text:`⚠️ วันนี้บันทึกไปแล้ว ${already} ชม.` });

    const after    = text.replace(/#OT/i,"").replace(startTime,"").replace(endTime,"").trim();
    const [task="", location=""] = after.includes("|") ? after.split("|").map(s=>s.trim()) : [after, ""];
    const pay      = Math.round(hours * empData.hourlyRate * WEEKDAY_MULTIPLIER);
    await saveRecord(sheets, { name:senderName, date:todayDate, startTime, endTime, hours, task:task||"-", location, otType:"วันธรรมดา", pay });
    return client.replyMessage(event.replyToken, { type:"text", text:`✅ บันทึก OT\n👤 ${senderName}\n⏰ ${startTime}–${endTime} (${hours}ชม.)\n📝 ${task||"-"}\n📅 ${todayDate}` });

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
app.listen(PORT, () => console.log(`🟢 OT Bot + LIFF on port ${PORT}`));
