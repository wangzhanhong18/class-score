import { updateStudentCalculations } from './stats.js';

let toastTimer = null;
function showToast(msg) {
  const el = document.getElementById('toast');
  if (!el) return;
  clearTimeout(toastTimer);
  el.textContent = msg;
  el.classList.add('show');
  toastTimer = setTimeout(() => el.classList.remove('show'), 2500);
}

/**
 * Render an editable score table with all exam columns.
 * Columns: 座位号 | 姓名 | [subjects] | 总分 | 班排名 | 级排名
 */
export function renderTable(container, students, subjects, onChange) {
  container.innerHTML = '';

  updateStudentCalculations(students);

  const wrapper = document.createElement('div');
  wrapper.className = 'table-scroll';

  const table = document.createElement('table');
  table.className = 'score-table';

  /* THEAD */
  const thead = document.createElement('thead');
  const tr = document.createElement('tr');
  const appendTh = (text, cls) => {
    const th = document.createElement('th');
    th.textContent = text;
    th.className = cls || '';
    tr.appendChild(th);
  };
  appendTh('座位号', 'col-seat');
  appendTh('姓名', 'col-name');
  subjects.forEach(s => appendTh(s, 'col-subject'));
  appendTh('总分', 'col-calc');
  appendTh('班排名', 'col-rank');
  appendTh('级排名', 'col-rank');
  thead.appendChild(tr);
  table.appendChild(thead);

  /* TBODY */
  const tbody = document.createElement('tbody');

  students.forEach((student, rowIdx) => {
    const row = document.createElement('tr');

    // 座位号 (read-only)
    const seatTd = document.createElement('td');
    seatTd.className = 'cell-seat';
    seatTd.textContent = student.seatNo !== null && student.seatNo !== undefined ? String(student.seatNo) : '';
    row.appendChild(seatTd);

    // 姓名 (read-only)
    const nameTd = document.createElement('td');
    nameTd.className = 'cell-name';
    nameTd.textContent = student.name;
    row.appendChild(nameTd);

    // Score cells (editable)
    subjects.forEach(sub => {
      const td = document.createElement('td');
      td.className = 'cell-score';
      td.contentEditable = 'true';
      td.dataset.row = rowIdx;
      td.dataset.sub = sub;
      const val = student.scores[sub];
      td.textContent = val !== null ? String(val) : '';
      applyCellStyle(td, val);

      td.addEventListener('focus', onCellFocus);
      td.addEventListener('blur', () => onCellBlur(td, student, sub, students, table, onChange));
      td.addEventListener('keydown', e => onCellKeydown(e, td, rowIdx, subjects, table));
      row.appendChild(td);
    });

    // 总分 (calculated)
    const totalTd = document.createElement('td');
    totalTd.className = 'cell-calc';
    totalTd.id = `total-${rowIdx}`;
    totalTd.textContent = student.total;
    row.appendChild(totalTd);

    // 班排名 (calculated)
    const classRankTd = document.createElement('td');
    classRankTd.className = 'cell-rank';
    classRankTd.id = `classrank-${rowIdx}`;
    classRankTd.textContent = student.rank;
    row.appendChild(classRankTd);

    // 级排名 (read-only from original data)
    const gradeRankTd = document.createElement('td');
    gradeRankTd.className = 'cell-rank';
    gradeRankTd.id = `graderank-${rowIdx}`;
    gradeRankTd.textContent = student.gradeRank !== null && student.gradeRank !== undefined ? String(student.gradeRank) : '-';
    row.appendChild(gradeRankTd);

    tbody.appendChild(row);
  });

  table.appendChild(tbody);
  wrapper.appendChild(table);
  container.appendChild(wrapper);
}

/* Cell helpers */
function applyCellStyle(td, value) {
  td.classList.remove('cell-empty', 'cell-excellent', 'cell-fail');
  if (value === null || value === undefined) {
    td.classList.add('cell-empty');
  } else if (value >= 90) {
    td.classList.add('cell-excellent');
  } else if (value < 60) {
    td.classList.add('cell-fail');
  }
}

function onCellFocus(e) {
  const td = e.currentTarget;
  td.dataset.original = td.textContent;
  const range = document.createRange();
  range.selectNodeContents(td);
  const sel = window.getSelection();
  sel.removeAllRanges();
  sel.addRange(range);
}

function onCellBlur(td, student, sub, students, table, onChange) {
  const raw = td.textContent.trim();
  const prev = td.dataset.original || '';
  if (raw === '' && prev === '') return;
  if (raw === prev) return;

  if (raw === '') {
    student.scores[sub] = null;
  } else {
    const num = Number(raw);
    if (isNaN(num) || !isFinite(num) || num < 0 || num > 100) {
      td.textContent = prev;
      showToast('请输入 0～100 之间的有效分数');
      return;
    }
    student.scores[sub] = Math.round(num * 10) / 10;
  }

  td.textContent = student.scores[sub] !== null ? String(student.scores[sub]) : '';
  applyCellStyle(td, student.scores[sub]);

  updateStudentCalculations(students);
  refreshCalcCells(students, table);
  if (onChange) onChange(students);
}

function onCellKeydown(e, td, rowIdx, subjects, table) {
  if (e.key === 'Enter') {
    e.preventDefault();
    td.blur();
    focusCell(rowIdx + 1, getColIndex(td, subjects), table);
    return;
  }
  if (e.key === 'Tab') {
    e.preventDefault();
    td.blur();
    const col = getColIndex(td, subjects);
    const dir = e.shiftKey ? -1 : 1;
    let nextRow = rowIdx;
    let nextCol = col + dir;
    if (nextCol < 0) { nextCol = subjects.length - 1; nextRow--; }
    if (nextCol >= subjects.length) { nextCol = 0; nextRow++; }
    if (nextRow >= 0) focusCell(nextRow, nextCol, table);
    return;
  }
  if (e.key === 'Escape') {
    e.preventDefault();
    td.textContent = td.dataset.original || '';
    td.blur();
    return;
  }
  if (e.key === 'ArrowDown') {
    e.preventDefault();
    td.blur();
    focusCell(rowIdx + 1, getColIndex(td, subjects), table);
    return;
  }
  if (e.key === 'ArrowUp') {
    e.preventDefault();
    td.blur();
    focusCell(rowIdx - 1, getColIndex(td, subjects), table);
    return;
  }
  if (e.key === 'ArrowLeft' || e.key === 'ArrowRight') {
    const sel = window.getSelection();
    if (sel.rangeCount === 0) return;
    const r = sel.getRangeAt(0);
    if ((e.key === 'ArrowLeft' && r.startOffset === 0 && r.collapsed) ||
        (e.key === 'ArrowRight' && r.startOffset === td.textContent.length && r.collapsed)) {
      e.preventDefault();
      td.blur();
      const col = getColIndex(td, subjects);
      const dir = e.key === 'ArrowLeft' ? -1 : 1;
      const nextCol = col + dir;
      if (nextCol >= 0 && nextCol < subjects.length) focusCell(rowIdx, nextCol, table);
    }
    return;
  }
}

function getColIndex(td, subjects) {
  return subjects.indexOf(td.dataset.sub);
}

function focusCell(rowIdx, colIdx, table) {
  const rows = table.querySelectorAll('tbody tr');
  if (rowIdx < 0 || rowIdx >= rows.length) return;
  const cells = rows[rowIdx].querySelectorAll('td.cell-score');
  if (colIdx < 0 || colIdx >= cells.length) return;
  cells[colIdx].focus();
}

function refreshCalcCells(students, table) {
    students.forEach((s, i) => {
      const totalEl = table.querySelector(`#total-${i}`);
      const rankEl = table.querySelector(`#classrank-${i}`);
      const gradeRankEl = table.querySelector(`#graderank-${i}`);
      if (totalEl) totalEl.textContent = s.total;
      if (rankEl) rankEl.textContent = s.rank;
      if (gradeRankEl) gradeRankEl.textContent = s.gradeRank !== null && s.gradeRank !== undefined ? String(s.gradeRank) : '-';
    });
  }
