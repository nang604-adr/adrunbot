// ============================================================
// OT LINE Bot — index.js
// บริษัท Adrun | Node.js + LINE Messaging API + Google Sheets
// Features: บันทึก OT, วันหยุด, สรุปรายเดือน, ตรวจ limit,
//           คำนวณค่า OT อัตโนมัติ, แจ้งเตือนเกิน limit
// ============================================================
require("dotenv").config();
const express    = require("express");
const line       = require("@line/bot-sdk");
const { google } = require("googleapis");

// ── LINE Config ──────────────────────────────────────────────
const lineConfig = {
  channelAccessToken: process.env.LINE_TOKEN,
  channelSecret:      process.env.LINE_SECRET,
};
const client = new line.Client(lineConfig);
const app    = express();

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
const MAX_OT_PER_DAY     = 5;    // ชม. สูงสุดต่อวัน (จ-ส)
const WEEKDAY_MULTIPLIER = 1.5;  // คูณค่าแรงวันธรรมดา

// ── Webhook endpoint ─────────────────────────────────────────
app.post("/webhook", async (req, res) => {
  const sig = req.headers["x-line-signature"];
  if (!sig) return res.json({ status: "ok" });
  next();
}, line.middleware(lineConfig), async (req, res) => {
  res.json({ status: "ok" }); // ตอบ LINE ก่อนเสมอ
  await Promise.all(req.body.events.map(handleEvent));
});

// ── Health check ──────────────────────────────────────────────
app.get("/", (_, res) => res.send("🟢 OT Adrun Bot is running"));

// ── Main Event Handler ────────────────────────────────────────
async function handleEvent(event) {
  if (event.type !== "message" || event.message.type !== "text") return;

  const text       = event.message.text.trim();
  const lower      = text.toLowerCase();
  const replyToken = event.replyToken;

  // รับเฉพาะคำสั่งที่ขึ้นต้นด้วย #OT หรือ #โอที
  if (!lower.startsWith("#ot") && !lower.startsWith("#โอที")) return;

  try {
    // ──── ดึงชื่อผู้ส่ง ────
    const groupId = event.source.groupId;
    const userId  = event.source.userId;
    let senderName = "ไม่ทราบชื่อ";
    try {
      if (groupId) {
        const profile = await client.getGroupMemberProfile(groupId, userId);
        senderName = profile.displayName;
      } else {
        const profile = await client.getProfile(userId);
        senderName = profile.displayName;
      }
    } catch (_) {}

    const sheets = await getSheetsClient();

    // ══════════════════════════════════════════════
    // คำสั่ง: #OT สรุป
    // ══════════════════════════════════════════════
    if (lower.includes("สรุป")) {
      const msg = await buildSummaryMessage(sheets, senderName);
      return client.replyMessage(replyToken, msg);
    }

    // ══════════════════════════════════════════════
    // คำสั่ง: #OT ช่วย / #OT help
    // ══════════════════════════════════════════════
    if (lower.includes("ช่วย") || lower.includes("help") || lower.includes("วิธี")) {
      return client.replyMessage(replyToken, helpMessage());
    }

    // ──── ดึง Employees จาก Sheet ────
    const employees = await getEmployees(sheets);
    const empData   = employees.find(e => e.name === senderName);

    if (!empData) {
      return client.replyMessage(replyToken, txt(
        `⚠️ ไม่พบชื่อ "${senderName}" ในระบบ\n\n` +
        `กรุณาแจ้ง Admin เพื่อเพิ่มชื่อพนักงาน\n` +
        `(ชื่อต้องตรงกับชื่อไลน์ของคุณ 100%)`
      ));
    }

    // ──── ตรวจวันหยุด ────
    const todayDate = getTodayThai();
    const todayDow  = new Date().getDay(); // 0=อาทิตย์, 6=เสาร์
    const holidays  = await getHolidays(sheets);
    const isHolidayDate = holidays.includes(todayDate);
    const isSunday  = todayDow === 0;

    const isHolidayCmd = lower.includes("วันหยุด") || lower.includes("หยุด");

    // ══════════════════════════════════════════════
    // คำสั่ง: #OT วันหยุด [งาน] [สถานที่]
    // ══════════════════════════════════════════════
    if (isHolidayCmd || isHolidayDate || isSunday) {
      // parse งานและสถานที่ออกจากข้อความ
      const parts    = text.replace(/#OT/i, "").replace(/#โอที/i, "")
                           .replace(/วันหยุด/g, "").replace(/หยุด/g, "")
                           .trim().split(/\s{2,}|\t/); // แยกด้วย 2 space หรือ tab
      const allWords = text.replace(/#OT/i, "").replace(/#โอที/i, "")
                           .replace(/วันหยุด/g, "").replace(/หยุด/g, "")
                           .trim();
      // งาน = ก่อน | , สถานที่ = หลัง |
      let task = "", location = "";
      if (allWords.includes("|")) {
        [task, location] = allWords.split("|").map(s => s.trim());
      } else {
        task = allWords;
      }

      const typeLabel = isSunday ? "วันอาทิตย์" : isHolidayDate ? "วันหยุดนักขัตฤกษ์" : "วันหยุด";

      await saveRecord(sheets, {
        name: senderName, date: todayDate,
        startTime: "-", endTime: "-", hours: 0,
        task: task || "-", location,
        otType: typeLabel, pay: empData.holidayFlat,
      });

      return client.replyMessage(replyToken, txt(
        `✅ บันทึก OT สำเร็จ!\n` +
        `👤 ${senderName}\n` +
        `🌅 ${typeLabel}\n` +
        `📝 ${task || "-"}` +
        (location ? `\n📍 ${location}` : "") +
        `\n📅 ${todayDate}`
      ));
    }

    // ══════════════════════════════════════════════
    // คำสั่ง: #OT [เวลาเริ่ม] [เวลาสิ้นสุด] [งาน] | [สถานที่]
    // ตัวอย่าง: #OT 18:00 21:00 ซ่อมสายพาน | โรงงาน A
    // ══════════════════════════════════════════════
    const timePattern = /\b(\d{1,2}):(\d{2})\b/g;
    const times       = [...text.matchAll(timePattern)];

    if (times.length < 2) {
      return client.replyMessage(replyToken, txt(
        `❓ ไม่เข้าใจคำสั่ง\n\n` +
        `รูปแบบที่ถูกต้อง:\n` +
        `#OT 18:00 21:00 [งาน] | [สถานที่]\n\n` +
        `พิมพ์ #OT ช่วย เพื่อดูคำสั่งทั้งหมด`
      ));
    }

    const startTime = times[0][0];
    const endTime   = times[1][0];
    const hours     = calcHours(startTime, endTime);

    if (hours <= 0) {
      return client.replyMessage(replyToken, txt(
        `⚠️ เวลาสิ้นสุดต้องมากกว่าเวลาเริ่มต้น\n` +
        `(เริ่ม ${startTime}, สิ้นสุด ${endTime})`
      ));
    }

    if (hours > MAX_OT_PER_DAY) {
      return client.replyMessage(replyToken, txt(
        `⚠️ OT วันธรรมดาทำได้สูงสุด ${MAX_OT_PER_DAY} ชม./วัน\n` +
        `คุณระบุ ${hours} ชม. — กรุณาแก้ไข`
      ));
    }

    // ตรวจว่าวันนี้บันทึกไปแล้วกี่ชม.
    const alreadyToday = await getTodayHours(sheets, senderName, todayDate);
    const totalIfAdd   = alreadyToday + hours;

    if (totalIfAdd > MAX_OT_PER_DAY) {
      const remain = +(MAX_OT_PER_DAY - alreadyToday).toFixed(1);
      return client.replyMessage(replyToken, txt(
        `⚠️ วันนี้ ${senderName} บันทึก OT ไปแล้ว ${alreadyToday} ชม.\n` +
        `เพิ่มได้อีกสูงสุด ${remain} ชม. (รวมไม่เกิน ${MAX_OT_PER_DAY} ชม./วัน)`
      ));
    }

    // parse งาน | สถานที่
    const afterTimes = text
      .replace(/#OT/i, "").replace(/#โอที/i, "")
      .replace(startTime, "").replace(endTime, "")
      .trim();

    let task = "", location = "";
    if (afterTimes.includes("|")) {
      [task, location] = afterTimes.split("|").map(s => s.trim());
    } else {
      task = afterTimes;
    }

    const pay         = Math.round(hours * empData.hourlyRate * WEEKDAY_MULTIPLIER);
    const dayNamesTH  = ["อาทิตย์","จันทร์","อังคาร","พุธ","พฤหัสบดี","ศุกร์","เสาร์"];
    const dayName     = dayNamesTH[todayDow];

    await saveRecord(sheets, {
      name: senderName, date: todayDate,
      startTime, endTime, hours,
      task: task || "-", location,
      otType: "วันธรรมดา", pay,
    });

    // แจ้งเตือนถ้าใกล้ถึง limit
    const warningLine = totalIfAdd >= MAX_OT_PER_DAY
      ? `\n⚠️ ครบ ${MAX_OT_PER_DAY} ชม. แล้ววันนี้`
      : totalIfAdd >= MAX_OT_PER_DAY - 1
      ? `\n📢 วันนี้รวม ${totalIfAdd} ชม. (เหลืออีก ${+(MAX_OT_PER_DAY - totalIfAdd).toFixed(1)} ชม.)`
      : "";

    return client.replyMessage(replyToken, txt(
      `✅ บันทึก OT สำเร็จ!\n` +
      `👤 ${senderName}\n` +
      `⏰ ${startTime}–${endTime} (${hours} ชม.) วัน${dayName}\n` +
      `📝 ${task || "-"}` +
      (location ? `\n📍 ${location}` : "") +
      `\n📅 ${todayDate}` +
      warningLine
    ));

  } catch (err) {
    console.error("[OT Bot Error]", err.message || err);
    return client.replyMessage(event.replyToken, txt(
      "❌ เกิดข้อผิดพลาด กรุณาลองใหม่\nหรือแจ้ง Admin หากปัญหายังคงอยู่"
    ));
  }
}

// ══════════════════════════════════════════════════════════════
// Helper Functions
// ══════════════════════════════════════════════════════════════

function txt(text) {
  return { type: "text", text };
}

function helpMessage() {
  return txt(
    `📖 คำสั่ง OT Bot\n\n` +
    `📅 วันธรรมดา (จ–ส):\n` +
    `#OT 18:00 21:00 [งาน] | [สถานที่]\n` +
    `ตย: #OT 18:00 21:00 ซ่อมเครื่อง | โรงงาน A\n\n` +
    `🌅 วันหยุด/อาทิตย์:\n` +
    `#OT วันหยุด [งาน] | [สถานที่]\n` +
    `ตย: #OT วันหยุด ทำรายงาน | ออฟฟิศ\n\n` +
    `📊 ดูสรุปของตัวเอง:\n#OT สรุป\n\n` +
    `⏱ OT วันธรรมดา สูงสุด ${MAX_OT_PER_DAY} ชม./วัน\n` +
    `(ถ้าวันนั้นเป็นวันหยุด ระบบตรวจให้อัตโนมัติ)`
  );
}

function getTodayThai() {
  const d  = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Bangkok" }));
  const dd = String(d.getDate()).padStart(2, "0");
  const mm = String(d.getMonth() + 1).padStart(2, "0");
  const yy = d.getFullYear() + 543;
  return `${dd}/${mm}/${yy}`;
}

function calcHours(start, end) {
  const [sh, sm] = start.split(":").map(Number);
  const [eh, em] = end.split(":").map(Number);
  const mins = eh * 60 + em - sh * 60 - sm;
  return mins > 0 ? +(mins / 60).toFixed(2) : 0;
}

// ══════════════════════════════════════════════════════════════
// Google Sheets Functions
// ══════════════════════════════════════════════════════════════

async function getEmployees(sheets) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "Employees!A2:C200",
  });
  return (res.data.values || []).map(row => ({
    name:        (row[0] || "").trim(),
    hourlyRate:  Number(row[1]) || 80,
    holidayFlat: Number(row[2]) || 500,
  })).filter(e => e.name);
}

async function getHolidays(sheets) {
  try {
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: SHEET_ID,
      range: "Holidays!A2:A100",
    });
    return (res.data.values || []).map(r => (r[0] || "").trim());
  } catch (_) {
    return [];
  }
}

async function saveRecord(sheets, data) {
  const now = new Date().toLocaleString("th-TH", { timeZone: "Asia/Bangkok" });
  await sheets.spreadsheets.values.append({
    spreadsheetId: SHEET_ID,
    range: "OT_Records!A:K",
    valueInputOption: "USER_ENTERED",
    resource: {
      values: [[
        data.name,
        data.date,
        data.startTime,
        data.endTime,
        data.hours,
        data.task,
        data.location || "",
        data.otType,
        data.pay,
        now,
      ]],
    },
  });
}

async function getTodayHours(sheets, name, date) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "OT_Records!A:H",
  });
  return (res.data.values || [])
    .filter(r => r[0] === name && r[1] === date && r[7] === "วันธรรมดา")
    .reduce((sum, r) => sum + (Number(r[4]) || 0), 0);
}

async function buildSummaryMessage(sheets, name) {
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SHEET_ID,
    range: "OT_Records!A:I",
  });

  const now   = new Date(new Date().toLocaleString("en-US", { timeZone: "Asia/Bangkok" }));
  const mm    = String(now.getMonth() + 1).padStart(2, "0");
  const yyyy  = String(now.getFullYear() + 543);
  const month = `${mm}/${yyyy}`;

  const rows = (res.data.values || []).slice(1);
  const mine = rows.filter(r => r[0] === name && r[1] && r[1].endsWith(`/${mm}/${yyyy}`));

  if (mine.length === 0) {
    return txt(`📊 ${name}\nยังไม่มี OT เดือน ${mm}/${yyyy}`);
  }

  const totalH   = mine.filter(r => r[7] === "วันธรรมดา").reduce((s, r) => s + (Number(r[4]) || 0), 0);
  const holidays = mine.filter(r => r[7] !== "วันธรรมดา").length;
  const last5    = mine.slice(-5).map(r => {
    const isHol = r[7] !== "วันธรรมดา";
    return `📅 ${r[1]}  ${isHol ? "🌅 " + r[7] : `⏰ ${r[2]}–${r[3]} (${r[4]}ชม.)`}  📝 ${r[5] || "-"}`;
  });

  return txt(
    `📊 สรุป OT ของ ${name}\nเดือน ${mm}/${yyyy}\n\n` +
    `⏱ วันธรรมดา: ${totalH} ชม.\n` +
    `🌅 วันหยุด: ${holidays} วัน\n\n` +
    `รายการล่าสุด:\n${last5.join("\n")}`
  );
}

// ── Start ─────────────────────────────────────────────────────
const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🟢 OT Bot ready on port ${PORT}`));
