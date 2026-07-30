/**
 * Recalculate total, average, and rank for each student (mutates in place).
 * Rank: descending by total; same total broken by seatNo (smaller = higher rank).
 */
export function updateStudentCalculations(students) {
  students.forEach(s => {
    const entries = Object.values(s.scores).filter(v => v !== null && !isNaN(v));
    const total = entries.reduce((sum, v) => sum + v, 0);
    s.total = total;
    s.average = entries.length > 0 ? Math.round((total / entries.length) * 10) / 10 : 0;
  });

  const sorted = [...students].sort((a, b) => {
    if (b.total !== a.total) return b.total - a.total;
    return (a.seatNo || 99999) - (b.seatNo || 99999);
  });
  sorted.forEach((s, i) => { s.rank = i + 1; });
}

/**
 * Compute per-subject statistics and overall summary.
 */
/**
 * Calculate grade-level rankings across all sheets.
 * Updates student.gradeRank in place.
 * Same total broken by seatNo (smaller = higher rank).
 */
export function calculateGradeRank(sheets) {
  const all = [];
  sheets.forEach(sheet => sheet.students.forEach(s => all.push({ student: s })));
  all.sort((a, b) => {
    if (b.student.total !== a.student.total) return b.student.total - a.student.total;
    return (a.student.seatNo || 99999) - (b.student.seatNo || 99999);
  });
  all.forEach((item, i) => { item.student.gradeRank = i + 1; });
}

export function calculateStats(students, subjects) {
  const overall = {
    totalStudents: students.length,
    completeCount: students.filter(s =>
      subjects.every(sub => s.scores[sub] !== null && !isNaN(s.scores[sub]))
    ).length,
  };

  const subjectStats = {};
  subjects.forEach(sub => {
    const vals = students
      .map(s => s.scores[sub])
      .filter(v => v !== null && !isNaN(v));

    if (vals.length === 0) {
      subjectStats[sub] = {
        max: "-", min: "-", avg: "-",
        excellent: 0, excellentRate: 0,
        pass: 0, passRate: 0,
        fail: 0, failRate: 0,
        count: 0,
      };
      return;
    }

    const max = Math.max(...vals);
    const min = Math.min(...vals);
    const avg = Math.round((vals.reduce((a, b) => a + b, 0) / vals.length) * 10) / 10;
    const excellent = vals.filter(v => v >= 90).length;
    const pass = vals.filter(v => v >= 60).length;
    const fail = vals.filter(v => v < 60).length;

    subjectStats[sub] = {
      max, min, avg,
      excellent,
      excellentRate: Math.round((excellent / vals.length) * 100),
      pass,
      passRate: Math.round((pass / vals.length) * 100),
      fail,
      failRate: Math.round((fail / vals.length) * 100),
      count: vals.length,
    };
  });

  return { overall, subjectStats };
}
