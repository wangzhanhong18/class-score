const NAME_PATTERNS = ["姓名", "学生姓名", "学生", "name"];
const SKIP_PATTERNS = ["总分", "平均分", "排名", "total", "average", "rank", "sum", "avg"];
const SUBJECT_PATTERNS = [
  "语文", "数学", "英语", "物理", "化学", "生物", "历史", "地理",
  "政治", "科学", "道德与法治", "体育", "音乐", "美术", "信息技术",
  "综合实践", "劳动", "外语", "日语", "俄语", "法语", "德语",
  "chinese", "math", "english", "physics", "chemistry", "biology",
  "history", "geography", "science", "music", "art", "pe",
];

function findHeaderRow(data) {
  for (let r = 0; r < Math.min(data.length, 6); r++) {
    const row = data[r];
    if (!row || !Array.isArray(row) || row.length < 2) continue;
    const headers = row.map(h => String(h).trim());
    if (headers.some(h => h === "座位号" || h === "座号") && headers.some(h => h === "姓名")) return r;
  }
  return -1;
}

function isExamFormat(data) {
  return findHeaderRow(data) >= 0;
}

function parseExamSheet(data) {
  if (data.length < 2) return null;
  const headerRowIdx = findHeaderRow(data);
  if (headerRowIdx < 0) return null;

  const headers = data[headerRowIdx].map(h => String(h).trim());
  let seatIdx = -1, nameIdx = -1;
  headers.forEach((h, i) => {
    if (h === "座位号" || h === "座号") seatIdx = i;
    if (h === "姓名") nameIdx = i;
  });
  if (seatIdx === -1 || nameIdx === -1) return null;

  const calcHeaders = ["总分", "总成绩", "合计", "total", "sum", "平均分", "avg", "average", "班排名", "级排名", "排名", "rank", "班别"];
  const calcIdxs = new Set();
  headers.forEach((h, i) => {
    if (i === seatIdx || i === nameIdx) return;
    const lower = h.toLowerCase().replace(/\s+/g, "");
    if (calcHeaders.some(c => lower.includes(c))) calcIdxs.add(i);
  });

  // Identify grade rank column specifically
  let gradeRankIdx = -1;
  headers.forEach((h, i) => {
    const lower = h.toLowerCase().replace(/\s+/g, "");
    if (lower.includes("级排名") || lower.includes("年级排名")) gradeRankIdx = i;
  });

  const subjects = [], subjectIdxs = [];
  headers.forEach((h, i) => {
    if (i === seatIdx || i === nameIdx) return;
    if (calcIdxs.has(i)) return;
    const lower = h.toLowerCase().replace(/\s+/g, "");
    const hasNumeric = data.slice(headerRowIdx + 1).some(row => {
      const v = row[i];
      return v !== "" && v !== undefined && v !== null && !isNaN(Number(v));
    });
    const isKnown = SUBJECT_PATTERNS.some(p => lower.includes(p));
    if (hasNumeric || isKnown) { subjects.push(h); subjectIdxs.push(i); }
  });

  const students = [];
  for (let r = headerRowIdx + 1; r < data.length; r++) {
    const row = data[r];
    if (!row || !row[nameIdx]) continue;
    const name = String(row[nameIdx] || "").trim();
    if (!name || name === "平均分" || name === "总分") continue;
    const rawSeat = row[seatIdx];
    const seatNo = rawSeat !== "" && rawSeat !== undefined && rawSeat !== null && !isNaN(Number(rawSeat)) ? Number(rawSeat) : null;
    let gradeRank = null;
    if (gradeRankIdx >= 0) {
      const v = row[gradeRankIdx];
      gradeRank = (v !== "" && v !== undefined && v !== null && !isNaN(Number(v))) ? Number(v) : null;
    }
    const scores = {};
    subjectIdxs.forEach((idx, i) => {
      const v = row[idx];
      scores[subjects[i]] = (v !== "" && v !== undefined && v !== null && !isNaN(Number(v))) ? Number(v) : null;
    });
    students.push({ seatNo, name, scores, gradeRank });
  }
  return { students, subjects };
}

/** Check if ANY sheet in the workbook matches exam format. */
function detectExamFormat(workbook) {
  for (const name of workbook.SheetNames) {
    const data = XLSX.utils.sheet_to_json(workbook.Sheets[name], { header: 1 });
    if (data && data.length >= 2 && isExamFormat(data)) return true;
  }
  return false;
}

export async function parseExcel(file) {
  const buffer = await file.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: "array" });

  if (detectExamFormat(workbook)) {
    const sheets = [];
    for (const sheetName of workbook.SheetNames) {
      const sheet = workbook.Sheets[sheetName];
      const data = XLSX.utils.sheet_to_json(sheet, { header: 1 });
      if (!data || data.length < 2) continue;
      const parsed = parseExamSheet(data);
      if (parsed && parsed.students.length > 0) {
        sheets.push({ name: sheetName.trim(), ...parsed });
      }
    }
    if (sheets.length === 0) throw new Error("未能从任何工作表中读取到学生数据");
    return { fileName: file.name, sheets, workbook, isExamFormat: true };
  }

  // ---- Generic format: find first sheet with data ---- //
  let sheetName = "", sheetData = null;
  for (const name of workbook.SheetNames) {
    const s = workbook.Sheets[name];
    const d = XLSX.utils.sheet_to_json(s, { header: 1, defval: "" });
    if (d.length >= 2) { sheetName = name; sheetData = d; break; }
  }
  if (!sheetData) throw new Error("文件至少需要包含表头和一行学生数据");

  const headers = sheetData[0].map(h => String(h).trim());
  const rows = sheetData.slice(1);
  let nameIdx = -1, nameCol = "";
  const subjects = [], subjectIdxs = [];

  headers.forEach((h, i) => {
    const lower = h.toLowerCase().replace(/\s+/g, "");
    if (NAME_PATTERNS.some(p => lower.includes(p.toLowerCase().replace(/\s+/g, "")))) { nameIdx = i; nameCol = h; return; }
    if (SKIP_PATTERNS.some(p => lower.includes(p))) return;
    const hasNumeric = rows.some(row => { const v = row[i]; return v !== "" && v !== undefined && v !== null && !isNaN(Number(v)); });
    const isKnownSubject = SUBJECT_PATTERNS.some(p => lower.includes(p.toLowerCase().replace(/\s+/g, "")));
    if (hasNumeric || isKnownSubject) { subjects.push(h); subjectIdxs.push(i); }
  });

  if (nameIdx === -1) throw new Error("未找到学生名列");
  if (subjects.length === 0) throw new Error("未找到科目成绩列");

  const students = rows.map(row => {
    const name = String(row[nameIdx] || "").trim();
    if (!name) return null;
    const scores = {};
    subjectIdxs.forEach((idx, i) => { const v = row[idx]; scores[subjects[i]] = v !== "" && v !== undefined && v !== null && !isNaN(Number(v)) ? Number(v) : null; });
    return { name, scores };
  }).filter(s => s !== null);

  if (students.length === 0) throw new Error("未能读取到任何学生数据");
  return { fileName: file.name, sheets: [{ name: sheetName, students, subjects }], workbook, isExamFormat: false, nameCol };
}

export function exportToExcel(data) {
  const { sheets, fileName, isExamFormat, workbook } = data;

  if (isExamFormat && workbook) {
    // Create new workbook from updated data instead of modifying original
    const wbNew = XLSX.utils.book_new();
    sheets.forEach(sheetData => {
      const raw = XLSX.utils.sheet_to_json(workbook.Sheets[sheetData.name], { header: 1 });
      if (raw.length < 2) return;
      let headerRowIdx = -1;
      for (let r = 0; r < Math.min(raw.length, 6); r++) {
        const row = raw[r];
        if (!row || !Array.isArray(row)) continue;
        const h = row.map(c => String(c).trim());
        if (h.some(x => x === "座位号" || x === "座号") && h.some(x => x === "姓名")) { headerRowIdx = r; break; }
      }
      if (headerRowIdx < 0) { XLSX.utils.book_append_sheet(wbNew, XLSX.utils.aoa_to_sheet(raw), sheetData.name); return; }
      const headers = raw[headerRowIdx].map(h => String(h).trim());
      let nameIdx = -1, totalIdx = -1, classRankIdx = -1, gradeRankIdx = -1;
      headers.forEach((h, i) => {
        if (h === "姓名") nameIdx = i;
        if (h.toLowerCase().includes("总分")) totalIdx = i;
        if (h === "班排名") classRankIdx = i;
        if (h === "级排名" || h === "年级排名") gradeRankIdx = i;
      });
      const students = sheetData.students;
      const sorted = [...students].sort((a, b) => {
        const ta = Object.values(a.scores).filter(v => v !== null && !isNaN(v)).reduce((s, v) => s + v, 0);
        const tb = Object.values(b.scores).filter(v => v !== null && !isNaN(v)).reduce((s, v) => s + v, 0);
        if (tb !== ta) return tb - ta;
        return (a.seatNo || 99999) - (b.seatNo || 99999);
      });
      const rankMap = new Map();
      sorted.forEach((s, i) => { rankMap.set(s.name, i + 1); });
      for (let r = headerRowIdx + 1; r < raw.length; r++) {
        const row = raw[r];
        if (!row) continue;
        const name = row[nameIdx] ? String(row[nameIdx]).trim() : "";
        const student = students.find(s => s.name === name);
        if (!student) continue;
        if (totalIdx >= 0) {
          const entries = Object.values(student.scores).filter(v => v !== null && !isNaN(v));
          row[totalIdx] = entries.reduce((sum, v) => sum + v, 0);
        }
        if (classRankIdx >= 0) row[classRankIdx] = rankMap.get(student.name) || "";
        if (gradeRankIdx >= 0 && student.gradeRank !== null && student.gradeRank !== undefined) row[gradeRankIdx] = student.gradeRank;
      }
      XLSX.utils.book_append_sheet(wbNew, XLSX.utils.aoa_to_sheet(raw), sheetData.name);
    });

    const out = XLSX.write(wbNew, { bookType: "xlsx", type: "array" });
    const baseName = fileName.replace(/\.(xlsx|xls)$/i, "");
    return { blob: new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), downloadName: baseName + "_已录入.xlsx" };
  }

  // Generic format: export ALL sheets as separate worksheets
  const wb = XLSX.utils.book_new();
  sheets.forEach((sheet, si) => {
    const { students, subjects, name } = sheet;
    const className = sheet.displayName || name;
    students.forEach(s => {
      const entries = Object.values(s.scores).filter(v => v !== null && !isNaN(v));
      s.total = entries.reduce((sum, v) => sum + v, 0);
      s.average = entries.length > 0 ? Math.round((s.total / entries.length) * 10) / 10 : 0;
    });
    const sorted = [...students].sort((a, b) => {
      if (b.total !== a.total) return b.total - a.total;
      return (a.seatNo || 99999) - (b.seatNo || 99999);
    });
    sorted.forEach((s, i) => { s.rank = i + 1; });
    const headers = ["班级", "座位号", "姓名", ...subjects, "总分", "平均分", "班排名", "级排名"];
    const body = students.map(s => [className, s.seatNo !== null && s.seatNo !== undefined ? String(s.seatNo) : "", s.name, ...subjects.map(sub => (s.scores[sub] !== null ? s.scores[sub] : "")), s.total, s.average, s.rank, s.gradeRank !== null && s.gradeRank !== undefined ? String(s.gradeRank) : ""]);
    const summary = ["平均分", "", ""];
    subjects.forEach(sub => { const vals = students.map(s => s.scores[sub]).filter(v => v !== null); summary.push(vals.length > 0 ? +(vals.reduce((a, b) => a + b, 0) / vals.length).toFixed(1) : ""); });
    summary.push("", "", "", "");
    const aoa = [headers, ...body, summary];
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    ws["!cols"] = [{ wch: 16 }, { wch: 8 }, { wch: 12 }, ...subjects.map(() => ({ wch: 10 })), { wch: 10 }, { wch: 10 }, { wch: 8 }, { wch: 8 }];
    XLSX.utils.book_append_sheet(wb, ws, name || ("成绩" + (si + 1)));
  });
  const out = XLSX.write(wb, { bookType: "xlsx", type: "array" });
  const baseName = fileName.replace(/\.(xlsx|xls)$/i, "");
  return { blob: new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }), downloadName: baseName + "_已录入.xlsx" };
}

