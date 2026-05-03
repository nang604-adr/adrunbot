// จำลอง logic ของ per-employee export เพื่อ verify
const XLSX = require("xlsx");

const recs = [
  { name: "ติน", date: "01/05/2569", startTime: "18:00", endTime: "22:00", hours: 4, task: "เซ็ตอุปกรณ์", location: "ออฟฟิศ", otType: "วันธรรมดา", pay: 320, paidAt: "PAY-TEST" },
  { name: "ติน", date: "05/05/2569", startTime: "-", endTime: "-", hours: 0, task: "ไป site เชียงใหม่", location: "เชียงใหม่", otType: "ต่างจังหวัด", pay: 600, paidAt: "PAY-TEST" },
  { name: "ติน", date: "06/05/2569", startTime: "-", endTime: "-", hours: 0, task: "ไป site เชียงใหม่", location: "เชียงใหม่", otType: "ต่างจังหวัด", pay: 600, paidAt: "PAY-TEST" },
  { name: "บอย", date: "02/05/2569", startTime: "08:00", endTime: "17:00", hours: 5, task: "ติดตั้งงาน", location: "ลูกค้า A", otType: "วันหยุด", pay: 500, paidAt: "PAY-TEST" },
];
const empListExp = [
  { name: "ติน", hourlyRate: 80, holidayFlat: 500, outProvinceFlat: 600, travelAllowance: 1500, socialSecurity: 750 },
  { name: "บอย", hourlyRate: 90, holidayFlat: 500, outProvinceFlat: 0, travelAllowance: 0, socialSecurity: 750 },
];
const empMapExp = {};
empListExp.forEach(e => empMapExp[e.name] = e);

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
  const ed = empMapExp[e.name] || {};
  const travel = ed.travelAllowance || 0;
  const social = ed.socialSecurity || 0;
  const netPay = e.pay + travel - social;
  return { name: e.name, days: e.days.size, hours: +e.hours.toFixed(2),
           holidays: e.holidays, pay: e.pay, travel, social, netPay };
}).sort((a,b) => b.netPay - a.netPay);

const wb = XLSX.utils.book_new();
const ws1 = XLSX.utils.aoa_to_sheet([["ทดสอบสรุป"]]);
XLSX.utils.book_append_sheet(wb, ws1, "สรุป");

const thaiDays = ["อา","จ","อ","พ","พฤ","ศ","ส"];
function thaiDayOfDate(s) { try { const [d,m,y]=s.split("/").map(Number); return thaiDays[new Date(y-543,m-1,d).getDay()]||""; } catch{return"";} }
function safeSheetName(name) { let s=name.replace(/[:\\\/?*\[\]]/g,"_").trim(); if(s.length>28)s=s.slice(0,28); return s||"พนักงาน"; }

const recsByName = {};
recs.forEach(r => (recsByName[r.name]=recsByName[r.name]||[]).push(r));

const usedSheetNames = new Set(["สรุป"]);
summary.forEach(emp => {
  const list = (recsByName[emp.name]||[]).slice().sort((a,b)=>a.date.localeCompare(b.date));
  const ed = empMapExp[emp.name] || {};
  const hourlyRate = ed.hourlyRate || 80;
  let opNight = 0;
  const rows = [
    [`ใบรายการ OT — ${emp.name}`],
    [`รอบจ่าย: 30/04/2569    Payroll ID: PAY-TEST`],
    [],
    ["วันที่","วัน","รายละเอียดงาน","สถานที่","เริ่ม","สิ้นสุด","ชม.","ประเภท","ค่า (฿)"],
  ];
  list.forEach(r => {
    const isOutProv = (r.otType||"").includes("ต่างจังหวัด");
    let startCell = r.startTime, endCell = r.endTime;
    if (isOutProv) { opNight += 1; startCell = `ตจว. คืนที่ ${opNight}`; endCell = ""; }
    rows.push([r.date, thaiDayOfDate(r.date), r.task||"", r.location||"", startCell, endCell, r.hours||0, r.otType, r.pay]);
  });
  const totalOT = list.reduce((s,r)=>s+(r.pay||0),0);
  const totalHours = list.reduce((s,r)=>s+(r.hours||0),0);
  rows.push([]);
  rows.push(["","","","","","รวม", +totalHours.toFixed(2),"", totalOT]);
  rows.push([]);
  rows.push(["─── สรุปจ่าย ───"]);
  rows.push(["+ ค่า OT","","","","","","","", emp.pay]);
  if (emp.travel) rows.push(["+ ค่าเดินทาง","","","","","","","", emp.travel]);
  if (emp.social) rows.push(["- ประกันสังคม","","","","","","","", -emp.social]);
  rows.push(["ทำจ่ายสุทธิ","","","","","","","", emp.netPay]);
  rows.push([]);
  rows.push([`Rate: ${hourlyRate} บาท/ชม.    วันหยุด/ตจว.: ${ed.holidayFlat||0} บาท    ตจว.: ${ed.outProvinceFlat||0} บาท`]);

  const wsEmp = XLSX.utils.aoa_to_sheet(rows);
  let sname = safeSheetName(emp.name); let n=2;
  while (usedSheetNames.has(sname)) { sname = safeSheetName(emp.name)+"_"+n; n+=1; }
  usedSheetNames.add(sname);
  XLSX.utils.book_append_sheet(wb, wsEmp, sname);
});

XLSX.writeFile(wb, "/tmp/payroll_test.xlsx");

// อ่านกลับ verify
const rb = XLSX.readFile("/tmp/payroll_test.xlsx");
console.log("Sheets:", rb.SheetNames);
rb.SheetNames.forEach(sn => {
  console.log("\n=== Sheet:", sn, "===");
  const data = XLSX.utils.sheet_to_json(rb.Sheets[sn], { header: 1 });
  data.forEach(row => console.log("  ", JSON.stringify(row)));
});
