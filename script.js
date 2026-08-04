let hskData = [];

// Khởi chạy khi nạp trang
window.addEventListener('DOMContentLoaded', () => {
  fetch('HSK1_flashcards.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      hskData = rawData.slice(1).map(row => ({
        hanzi: row[0] || '',
        pinyin: row[1] || '',
        meaning: row[2] || '',
        exampleVn: row[3] || '',
        examplePinyin: row[4] || ''
      })).filter(item => item.hanzi && item.meaning);

      if (hskData.length > 0) {
        initMultipleChoiceApp();
        initBuilderApp();
      }
    })
    .catch(err => console.error("Lỗi đọc file Excel:", err));
});

/* ==========================================
   CHUYỂN TAB ĐỘC LẬP
   ========================================== */
function switchTab(tabId, btnElement) {
  document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

  document.getElementById(tabId).classList.add('active');
  btnElement.classList.add('active');
}

/* ==========================================
   TAB 1: TRẮC NGHIỆM TỪ VỰNG
   ========================================== */
let mcList = [];
let mcIndex = 0;
let currentMcItem = null;
let mcWrongList = JSON.parse(localStorage.getItem('hsk1_mc_wrong') || '[]');
let isMcReviewMode = false;

function initMultipleChoiceApp() {
  const savedMcOrder = localStorage.getItem('hsk1_mc_order');
  const savedMcIndex = localStorage.getItem('hsk1_mc_index');

  if (savedMcOrder) {
    const indices = JSON.parse(savedMcOrder);
    mcList = indices.map(idx => hskData[idx]).filter(Boolean);
  } else {
    const indices = hskData.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_mc_order', JSON.stringify(indices));
    mcList = indices.map(idx => hskData[idx]);
  }

  if (savedMcIndex) {
    mcIndex = parseInt(savedMcIndex, 10);
    if (mcIndex >= mcList.length) mcIndex = 0;
  }

  updateMcWrongCountUI();
  loadMcQuestion();
}

function loadMcQuestion() {
  document.getElementById('mc-feedback').innerText = '';
  
  let activeList = isMcReviewMode ? mcWrongList : mcList;

  if (activeList.length === 0) {
    if (isMcReviewMode) {
      alert("M không có câu sai nào trong danh sách!");
      toggleMcReviewMode();
      return;
    }
  }

  if (mcIndex >= activeList.length) mcIndex = 0;

  currentMcItem = activeList[mcIndex];

  // Cập nhật giao diện Tiến Độ
  document.getElementById('mc-progress-text').innerText = `${mcIndex + 1}/${activeList.length}`;
  const pct = Math.round(((mcIndex + 1) / activeList.length) * 100);
  document.getElementById('mc-progress-bar').style.width = `${pct}%`;

  document.getElementById('mc-meaning').innerText = `"${currentMcItem.meaning}"`;

  // Tạo 2 lựa chọn sai ngẫu nhiên
  const wrongOptions = hskData
    .filter(item => item.hanzi !== currentMcItem.hanzi)
    .sort(() => Math.random() - 0.5)
    .slice(0, 2);

  const options = [currentMcItem, ...wrongOptions].sort(() => Math.random() - 0.5);

  const container = document.getElementById('mc-options');
  container.innerHTML = '';

  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'option-btn';
    btn.onclick = () => checkMultipleChoice(opt);
    btn.innerHTML = `
      <span class="option-hanzi">${opt.hanzi}</span>
      <span class="option-pinyin">${opt.pinyin}</span>
    `;
    container.appendChild(btn);
  });
}

function checkMultipleChoice(selected) {
  const feedback = document.getElementById('mc-feedback');

  if (selected.hanzi === currentMcItem.hanzi) {
    feedback.style.color = '#1f883d';
    feedback.innerText = '🎉 Chính xác!';
    if (typeof confetti === 'function') confetti({ particleCount: 70, spread: 50, origin: { y: 0.4 } });

    // Nếu trả lời đúng ở chế độ ôn tập thì xóa khỏi danh sách sai
    if (isMcReviewMode) {
      mcWrongList = mcWrongList.filter(item => item.hanzi !== currentMcItem.hanzi);
      localStorage.setItem('hsk1_mc_wrong', JSON.stringify(mcWrongList));
      updateMcWrongCountUI();
    }

    nextMcQuestion();
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Chưa đúng, đã lưu vào danh sách sai!';

    // Thêm vào danh sách câu sai nếu chưa có
    if (!mcWrongList.some(item => item.hanzi === currentMcItem.hanzi)) {
      mcWrongList.push(currentMcItem);
      localStorage.setItem('hsk1_mc_wrong', JSON.stringify(mcWrongList));
      updateMcWrongCountUI();
    }
  }
}

function nextMcQuestion() {
  setTimeout(() => {
    mcIndex++;
    if (!isMcReviewMode) {
      localStorage.setItem('hsk1_mc_index', mcIndex);
    }
    loadMcQuestion();
  }, 1000);
}

function skipMcQuestion() {
  mcIndex++;
  if (!isMcReviewMode) localStorage.setItem('hsk1_mc_index', mcIndex);
  loadMcQuestion();
}

function updateMcWrongCountUI() {
  document.getElementById('mc-wrong-count').innerText = mcWrongList.length;
}

function toggleMcReviewMode() {
  if (!isMcReviewMode && mcWrongList.length === 0) {
    alert("Tuyệt vời! M chưa có câu sai nào!");
    return;
  }

  isMcReviewMode = !isMcReviewMode;
  mcIndex = 0;
  
  const badge = document.getElementById('mc-mode-badge');
  const btn = document.getElementById('mc-review-btn');

  if (isMcReviewMode) {
    badge.innerText = 'Chế độ: Ôn câu sai ⚠️';
    badge.style.background = '#fff8c5';
    badge.style.color = '#9a6700';
    btn.innerText = '↩️ Quay lại toàn bộ';
  } else {
    badge.innerText = 'Chế độ: Toàn bộ';
    badge.style.background = '#ddf4ff';
    badge.style.color = '#0969da';
    btn.innerHTML = `⚠️ Ôn câu sai (<span id="mc-wrong-count">${mcWrongList.length}</span>)`;
  }

  loadMcQuestion();
}

/* ==========================================
   TAB 2: SẮP XẾP CÂU HOÀN CHỈNH
   ========================================== */
let builderList = [];
let builderIndex = 0;
let currentCorrectWords = [];
let userSelectedPool = [];
let builderWrongList = JSON.parse(localStorage.getItem('hsk1_builder_wrong') || '[]');
let isBuilderReviewMode = false;

function initBuilderApp() {
  const validItems = hskData.filter(item => item.examplePinyin && item.exampleVn);

  const savedOrder = localStorage.getItem('hsk1_builder_order');
  const savedIndex = localStorage.getItem('hsk1_builder_index');

  if (savedOrder) {
    const indices = JSON.parse(savedOrder);
    builderList = indices.map(idx => validItems[idx]).filter(Boolean);
  } else {
    const indices = validItems.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_builder_order', JSON.stringify(indices));
    builderList = indices.map(idx => validItems[idx]);
  }

  if (savedIndex) {
    builderIndex = parseInt(savedIndex, 10);
    if (builderIndex >= builderList.length) builderIndex = 0;
  }

  updateBuilderWrongCountUI();
  loadBuilderQuestion();
}

function loadBuilderQuestion() {
  document.getElementById('builder-feedback').innerText = '';
  
  let activeList = isBuilderReviewMode ? builderWrongList : builderList;

  if (activeList.length === 0) {
    if (isBuilderReviewMode) {
      alert("M không có câu sai nào!");
      toggleBuilderReviewMode();
      return;
    }
  }

  if (builderIndex >= activeList.length) builderIndex = 0;

  const item = activeList[builderIndex];

  // Tiến độ
  document.getElementById('builder-progress-text').innerText = `${builderIndex + 1}/${activeList.length}`;
  const pct = Math.round(((builderIndex + 1) / activeList.length) * 100);
  document.getElementById('builder-progress-bar').style.width = `${pct}%`;

  document.getElementById('builder-meaning').innerText = `"${item.exampleVn}"`;

  let rawPinyin = item.examplePinyin.replace(/[。!？,.]/g, '').trim();
  currentCorrectWords = rawPinyin.split(/\s+/);

  let wordObjects = currentCorrectWords.map((word, id) => ({ id, text: word }));
  let shuffledWords = [...wordObjects].sort(() => Math.random() - 0.5);

  userSelectedPool = [];
  renderBuilderUI(shuffledWords);
}

function renderBuilderUI(poolWords) {
  const box = document.getElementById('selected-words-box');
  const pool = document.getElementById('word-pool');

  box.innerHTML = '';
  if (userSelectedPool.length === 0) {
    box.innerHTML = `<p class="placeholder-text">Click vào từng từ bên dưới để ghép câu...</p>`;
  } else {
    userSelectedPool.forEach((item, idx) => {
      const chip = document.createElement('span');
      chip.className = 'word-chip selected';
      chip.innerText = item.text;
      chip.onclick = () => unselectWord(idx, poolWords);
      box.appendChild(chip);
    });
  }

  pool.innerHTML = '';
  poolWords.forEach((item) => {
    const chip = document.createElement('button');
    chip.className = 'word-chip';
    chip.innerText = item.text;
    chip.onclick = () => selectWord(item, poolWords);
    pool.appendChild(chip);
  });
}

function selectWord(item, currentPool) {
  userSelectedPool.push(item);
  const newPool = currentPool.filter(w => w.id !== item.id);
  renderBuilderUI(newPool);
}

function unselectWord(selectedIdx, currentPool) {
  const removed = userSelectedPool.splice(selectedIdx, 1)[0];
  currentPool.push(removed);
  renderBuilderUI(currentPool);
}

function resetBuilderSelection() {
  loadBuilderQuestion();
}

function checkBuilderAnswer() {
  const feedback = document.getElementById('builder-feedback');
  const userResult = userSelectedPool.map(w => w.text).join(' ');
  const correctResult = currentCorrectWords.join(' ');

  if (userSelectedPool.length === 0) {
    feedback.style.color = '#d1242f';
    feedback.innerText = '⚠️ M chưa ghép từ nào kìa!';
    return;
  }

  const currentItem = (isBuilderReviewMode ? builderWrongList : builderList)[builderIndex];

  if (userResult.toLowerCase() === correctResult.toLowerCase()) {
    feedback.style.color = '#1f883d';
    feedback.innerText = `🎉 Chuẩn đét! Đáp án: "${correctResult}"`;
    if (typeof confetti === 'function') confetti({ particleCount: 80, spread: 60, origin: { y: 0.8 } });

    if (isBuilderReviewMode) {
      builderWrongList = builderWrongList.filter(item => item.exampleVn !== currentItem.exampleVn);
      localStorage.setItem('hsk1_builder_wrong', JSON.stringify(builderWrongList));
      updateBuilderWrongCountUI();
    }

    setTimeout(() => {
      builderIndex++;
      if (!isBuilderReviewMode) localStorage.setItem('hsk1_builder_index', builderIndex);
      loadBuilderQuestion();
    }, 1500);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Chưa đúng ngữ pháp rồi, thử lại xem!`;

    if (!builderWrongList.some(item => item.exampleVn === currentItem.exampleVn)) {
      builderWrongList.push(currentItem);
      localStorage.setItem('hsk1_builder_wrong', JSON.stringify(builderWrongList));
      updateBuilderWrongCountUI();
    }
  }
}

function skipBuilderSentence() {
  builderIndex++;
  if (!isBuilderReviewMode) localStorage.setItem('hsk1_builder_index', builderIndex);
  loadBuilderQuestion();
}

function updateBuilderWrongCountUI() {
  document.getElementById('builder-wrong-count').innerText = builderWrongList.length;
}

function toggleBuilderReviewMode() {
  if (!isBuilderReviewMode && builderWrongList.length === 0) {
    alert("M không có câu sai nào!");
    return;
  }

  isBuilderReviewMode = !isBuilderReviewMode;
  builderIndex = 0;

  const badge = document.getElementById('builder-mode-badge');
  const btn = document.getElementById('builder-review-btn');

  if (isBuilderReviewMode) {
    badge.innerText = 'Chế độ: Ôn câu sai ⚠️';
    badge.style.background = '#fff8c5';
    badge.style.color = '#9a6700';
    btn.innerText = '↩️ Quay lại toàn bộ';
  } else {
    badge.innerText = 'Chế độ: Toàn bộ';
    badge.style.background = '#ddf4ff';
    badge.style.color = '#0969da';
    btn.innerHTML = `⚠️ Ôn câu sai (<span id="builder-wrong-count">${builderWrongList.length}</span>)`;
  }

  loadBuilderQuestion();
}

/* Dark Mode */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
