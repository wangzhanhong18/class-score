import { parseExcel, exportToExcel } from './excel.js';
import { renderTable } from './table.js';
import { updateStudentCalculations, calculateStats, calculateGradeRank } from './stats.js';

export function initApp() {
  const fileInput = document.getElementById('fileInput');
  const exportBtn = document.getElementById('exportBtn');
  const tableContainer = document.getElementById('tableContainer');
  const emptyState = document.getElementById('emptyState');
  const statsPanel = document.getElementById('statsPanel');
  const fileInfo = document.getElementById('fileInfo');
  const classTabs = document.getElementById('classTabs');

  let appData = null;
  let currentSheetIndex = 0;

  function updateAllCalculations() {
    if (!appData) return;
    appData.sheets.forEach(sheet => updateStudentCalculations(sheet.students));
    calculateGradeRank(appData.sheets);
  }
 
  /* ---------- Upload ---------- */
  fileInput.addEventListener('change', async e => {
    const files = Array.from(e.target.files);
    if (files.length === 0) return;
    let allSheets = [];
    let errors = [];
    for (const file of files) {
      try {
        const parsed = await parseExcel(file);
        const baseName = file.name.replace(/\.(xlsx|xls)$/i, "");
        parsed.sheets.forEach(s => {
          s.displayName = files.length > 1 ? baseName + " - " + s.name : s.name;
        });
        allSheets = allSheets.concat(parsed.sheets);
      } catch (err) {
        errors.push(file.name + ": " + err.message);
      }
    }
    if (allSheets.length === 0) {
      alert("所有文件读取失败：\n" + errors.join("\n"));
      fileInput.value = "";
      return;
    }
    appData = {
      fileName: files.length > 1 ? "已合并 " + files.length + " 个文件" : files[0].name,
      sheets: allSheets,
      isExamFormat: false,
      workbook: null,
    };
    currentSheetIndex = 0;
    emptyState.style.display = 'none';
    classTabs.style.display = '';
    renderClassTabs();
    switchToClass(0);
    const totalStudents = allSheets.reduce((sum, s) => sum + s.students.length, 0);
    fileInfo.textContent = appData.fileName + " ｜ " + allSheets.length + " 个班级，共 " + totalStudents + " 名学生";
    exportBtn.disabled = false;
    if (errors.length > 0) {
      showToast("成功导入 " + (files.length - errors.length) + " 个文件，" + errors.length + " 个失败");
    }
    fileInput.value = "";
  });

  /* ---------- Export ---------- */
  exportBtn.addEventListener('click', () => {
      if (!appData) return;
      updateAllCalculations();
      const result = exportToExcel(appData);
    const url = URL.createObjectURL(result.blob);
    const a = document.createElement('a');
    a.href = url;
    a.download = result.downloadName;
    a.click();
    URL.revokeObjectURL(url);
  });

  /* ---------- Class Tabs ---------- */
  function renderClassTabs() {
    classTabs.innerHTML = '';
    classTabs.style.display = '';
    appData.sheets.forEach((sheet, i) => {
      const btn = document.createElement('button');
      btn.className = 'tab-btn' + (i === currentSheetIndex ? ' active' : '');
      btn.textContent = sheet.displayName || sheet.name;
      const count = document.createElement('span');
      count.className = 'tab-count';
      count.textContent = sheet.students.length + '人';
      btn.appendChild(count);
      btn.addEventListener('click', () => {
        currentSheetIndex = i;
        renderClassTabs();
        switchToClass(i);
      });
      classTabs.appendChild(btn);
    });
  }

  function switchToClass(index) {
      const sheet = appData.sheets[index];
      const { students, subjects } = sheet;
      updateAllCalculations();
      renderTable(tableContainer, students, subjects, onScoreChange);
    const stats = calculateStats(students, subjects);
    renderStats(stats, subjects);
    document.querySelectorAll('.tab-btn').forEach((btn, i) => {
      btn.classList.toggle('active', i === index);
    });
  }

  function onScoreChange(students) {
      updateAllCalculations();
      const sheet = appData.sheets[currentSheetIndex];
      const stats = calculateStats(students, sheet.subjects);
      renderStats(stats, sheet.subjects);
  }

  /* ---------- Create blank template ---------- */
  const createBlankBtn = document.getElementById('createBlankBtn');
  const createForm = document.getElementById('createForm');
  const createTemplateBtn = document.getElementById('createTemplateBtn');
  const namesInput = document.getElementById('namesInput');
  const subjectsInput = document.getElementById('subjectsInput');

  if (createBlankBtn) {
    createBlankBtn.addEventListener('click', () => {
      createForm.style.display = 'block';
      createBlankBtn.style.display = 'none';
    });
  }

  if (createTemplateBtn) {
    createTemplateBtn.addEventListener('click', () => {
      const nameList = namesInput.value.split('\n').map(n => n.trim()).filter(n => n.length > 0);
      const subjectList = subjectsInput.value.split(/[,，、\s]+/).map(s => s.trim()).filter(s => s.length > 0);
      if (nameList.length === 0) { showToast('请输入至少一个学生姓名'); return; }
      if (subjectList.length === 0) { showToast('请输入至少一个科目'); return; }

      const students = nameList.map(name => {
        const scores = {};
        subjectList.forEach(sub => { scores[sub] = null; });
        return { name, scores };
      });

      updateStudentCalculations(students);
      appData = {
        fileName: '空白模板.xlsx',
        sheets: [{ name: 'Sheet1', students, subjects: subjectList }],
        isExamFormat: false,
      };
      currentSheetIndex = 0;
      emptyState.style.display = 'none';
      classTabs.style.display = 'none';
      renderTable(tableContainer, students, subjectList, onScoreChange);
      const stats = calculateStats(students, subjectList);
      renderStats(stats, subjectList);
      exportBtn.disabled = false;
      fileInfo.textContent = `新建空白模板 | 1 个工作表，${students.length} 名学生，${subjectList.length} 个科目`;
    });
  }

  /* ---------- Reference Images ---------- */
  const imageInput = document.getElementById('imageInput');
  const imageGallery = document.getElementById('imageGallery');
  const galleryScroll = document.getElementById('galleryScroll');
  const lightbox = document.getElementById('lightbox');
  const lightboxImg = document.getElementById('lightboxImg');
  const lightboxCaption = document.getElementById('lightboxCaption');
  const lightboxClose = document.getElementById('lightboxClose');
  const refImages = [];

  imageInput.addEventListener('change', (e) => {
    const files = Array.from(e.target.files);
    files.forEach(file => {
      const url = URL.createObjectURL(file);
      refImages.push({ id: crypto.randomUUID(), name: file.name, url });
    });
    renderGallery();
    imageInput.value = '';
  });

  function renderGallery() {
    if (refImages.length === 0) { imageGallery.style.display = 'none'; return; }
    imageGallery.style.display = '';
    galleryScroll.innerHTML = '';
    refImages.forEach(img => {
      const thumb = document.createElement('div');
      thumb.className = 'gallery-thumb';
      thumb.innerHTML = '<img src="' + img.url + '" class="gallery-thumb-img">';
      const removeBtn = document.createElement('button');
      removeBtn.className = 'gallery-thumb-remove';
      removeBtn.innerHTML = '&times;';
      removeBtn.addEventListener('click', (e) => {
        e.stopPropagation();
        const idx = refImages.indexOf(img);
        if (idx !== -1) { URL.revokeObjectURL(img.url); refImages.splice(idx, 1); }
        renderGallery();
      });
      thumb.appendChild(removeBtn);
      thumb.addEventListener('click', () => {
        lightboxImg.src = img.url;
        lightboxCaption.textContent = img.name;
        lightbox.style.display = 'flex';
        document.body.style.overflow = 'hidden';
      });
      galleryScroll.appendChild(thumb);
    });
  }

  function closeLightbox() { lightbox.style.display = 'none'; document.body.style.overflow = ''; }
  lightboxClose.addEventListener('click', closeLightbox);
  lightbox.addEventListener('click', (e) => { if (e.target === lightbox) closeLightbox(); });
  document.addEventListener('keydown', (e) => { if (e.key === 'Escape') closeLightbox(); });

  /* ---------- Stats Panel ---------- */
  function renderStats(stats, subjects) {
    statsPanel.classList.add('active');
    let html = '<div class="stats-section"><div class="stats-section-title">概览</div><div class="stats-overview">';
    html += '<div class="stat-card"><div class="stat-value">' + stats.overall.totalStudents + '</div><div class="stat-label">学生总数</div></div>';
    html += '<div class="stat-card"><div class="stat-value">' + stats.overall.completeCount + '</div><div class="stat-label">已录入</div></div>';
    html += '</div></div>';
    html += '<div class="stats-section"><div class="stats-section-title">科目统计</div><div class="subject-stats">';
    subjects.forEach(sub => {
      const s = stats.subjectStats[sub];
      html += '<div class="subject-stat-card"><div class="subject-stat-header">' + sub + '</div>';
      html += '<div class="subject-stat-grid">';
      html += '<div class="subject-stat-item"><div class="s-value">' + s.max + '</div><div class="s-label">最高分</div></div>';
      html += '<div class="subject-stat-item"><div class="s-value">' + s.min + '</div><div class="s-label">最低分</div></div>';
      html += '<div class="subject-stat-item"><div class="s-value">' + s.avg + '</div><div class="s-label">平均分</div></div>';
      html += '</div>';
      html += '<div class="subject-stat-dist">';
      html += '<div class="dist-item dist-excellent"><div class="d-value">' + s.excellent + ' (' + s.excellentRate + '%)</div><div class="d-label">优秀 ≥90</div></div>';
      html += '<div class="dist-item dist-pass"><div class="d-value">' + s.pass + ' (' + s.passRate + '%)</div><div class="d-label">及格 ≥60</div></div>';
      html += '<div class="dist-item dist-fail"><div class="d-value">' + s.fail + ' (' + s.failRate + '%)</div><div class="d-label">不及格 <60</div></div>';
      html += '</div></div>';
    });
    html += '</div></div>';
    statsPanel.innerHTML = html;
  }

  let toastTimer = null;
  function showToast(msg) {
    const el = document.getElementById('toast');
    if (!el) return;
    clearTimeout(toastTimer);
    el.textContent = msg;
    el.classList.add('show');
    toastTimer = setTimeout(() => el.classList.remove('show'), 2500);
  }
}

