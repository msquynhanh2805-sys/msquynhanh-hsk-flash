let hskData = [];

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

function switchTab(tabId, btnElement) {
  document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

  document.getElementById(tabId).classList.add('active');
  btnElement.classList.add('active');
}

/* ==========================================
   TAB 1: TRẮC NGHIỆM TỪ VỰNG (CUSTOM SESSION)
   ========================================== */
let mcList = [];
let mcGlobalIndex = 0;
let mcSessionList = [];
let mcSessionIndex = 0;
let mcCorrectFirstTry = 0;
let mcWrongList = JSON.parse(localStorage.getItem('hsk1_mc_wrong') || '[]');
let isMcWrongMode = false;
let currentMcItem = null;

function initMultipleChoiceApp() {
  const savedOrder = localStorage.getItem('hsk1_mc_order');
  const savedIndex = localStorage.getItem('hsk1_mc_index');

  if (savedOrder) {
    const indices = JSON.parse(savedOrder);
    mcList = indices.map(idx => hskData[idx]).filter(Boolean);
  } else {
    const indices = hskData.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_mc_order', JSON.stringify(indices));
    mcList = indices.map(idx => hskData[idx]);
  }

  mcGlobalIndex = savedIndex ? parseInt(savedIndex, 10) : 0;
  if (mcGlobalIndex >= mcList.length) mcGlobalIndex = 0;

  updateMcSetupUI();
}

function updateMcSetupUI() {
  document.getElementById('mc-total-count').innerText = mcList.length;
  document.getElementById('mc-learned-count').innerText = mcGlobalIndex;
  document.getElementById('mc-remain-count').innerText = mcList.length - mcGlobalIndex;

  const wrongBtn = document.getElementById('mc-wrong-btn');
  if (mcWrongList.length > 0) {
    wrongBtn.style.display = 'inline-block';
    document.getElementById('mc-wrong-count').innerText = mcWrongList.length;
  } else {
    wrongBtn.style.display = 'none';
  }

  document.getElementById('mc-setup-screen').style.display = 'block';
  document.getElementById('mc-quiz-screen').style.display = 'none';
  document.getElementById('mc-summary-screen').style.display = 'none';
}

function startMcSession() {
  const countInput = parseInt(document.getElementById('mc-session-input').value, 10) || 10;
  
  if (mcGlobalIndex >= mcList.length) mcGlobalIndex = 0;

  mcSessionList = mcList.slice(mcGlobalIndex, mcGlobalIndex + countInput);
  if (mcSessionList.length === 0) {
    alert("Bạn đã hoàn thành toàn bộ kho từ vựng! Hệ thống sẽ reset lại từ đầu.");
    mcGlobalIndex = 0;
    localStorage.setItem('hsk1_mc_index', 0);
    mcSessionList = mcList.slice(0, countInput);
  }

  mcSessionIndex = 0;
  mcCorrectFirstTry = 0;
  isMcWrongMode = false;

  document.getElementById('mc-setup-screen').style.display = 'none';
  document.getElementById('mc-quiz-screen').style.display = 'block';
  document.getElementById('mc-mode-badge').innerText = 'Phiên học mới';

  loadMcQuestion();
}

function startMcWrongReview() {
  if (mcWrongList.length === 0) return;
  
  mcSessionList = [...mcWrongList];
  mcSessionIndex = 0;
  mcCorrectFirstTry = 0;
  isMcWrongMode = true;

  document.getElementById('mc-setup-screen').style.display = 'none';
  document.getElementById('mc-quiz-screen').style.display = 'block';
  document.getElementById('mc-mode-badge').innerText = 'Ôn câu sai ⚠️';

  loadMcQuestion();
}

function loadMcQuestion() {
  document.getElementById('mc-feedback').innerText = '';

  if (mcSessionIndex >= mcSessionList.length) {
    showMcSummary();
    return;
  }

  currentMcItem = mcSessionList[mcSessionIndex];

  document.getElementById('mc-progress-text').innerText = `${mcSessionIndex + 1}/${mcSessionList.length}`;
  const pct = Math.round(((mcSessionIndex + 1) / mcSessionList.length) * 100);
  document.getElementById('mc-progress-bar').style.width = `${pct}%`;

  document.getElementById('mc-meaning').innerText = `"${currentMcItem.meaning}"`;

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

let mcHadWrongInCurrent = false;

function checkMultipleChoice(selected) {
  const feedback = document.getElementById('mc-feedback');

  if (selected.hanzi === currentMcItem.hanzi) {
    feedback.style.color = '#1f883d';
    feedback.innerText = '🎉 Chính xác!';
    if (typeof confetti === 'function') confetti({ particleCount: 60, spread: 50, origin: { y: 0.4 } });

    if (!mcHadWrongInCurrent) mcCorrectFirstTry++;

    if (isMcWrongMode) {
      mcWrongList = mcWrongList.filter(item => item.hanzi !== currentMcItem.hanzi);
      localStorage.setItem('hsk1_mc_wrong', JSON.stringify(mcWrongList));
    }

    setTimeout(() => {
      mcSessionIndex++;
      if (!isMcWrongMode) {
        mcGlobalIndex++;
        localStorage.setItem('hsk1_mc_index', mcGlobalIndex);
      }
      mcHadWrongInCurrent = false;
      loadMcQuestion();
    }, 1000);

  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Chưa đúng, thử lại nhé!';
    mcHadWrongInCurrent = true;

    if (!mcWrongList.some(item => item.hanzi === currentMcItem.hanzi)) {
      mcWrongList.push(currentMcItem);
      localStorage.setItem('hsk1_mc_wrong', JSON.stringify(mcWrongList));
    }
  }
}

function skipMcQuestion() {
  mcSessionIndex++;
  if (!isMcWrongMode) {
    mcGlobalIndex++;
    localStorage.setItem('hsk1_mc_index', mcGlobalIndex);
  }
  mcHadWrongInCurrent = false;
  loadMcQuestion();
}

function showMcSummary() {
  document.getElementById('mc-quiz-screen').style.display = 'none';
  document.getElementById('mc-summary-screen').style.display = 'block';

  document.getElementById('mc-sum-done').innerText = mcSessionList.length;
  document.getElementById('mc-sum-correct').innerText = mcCorrectFirstTry;
  document.getElementById('mc-sum-wrong').innerText = mcSessionList.length - mcCorrectFirstTry;

  if (typeof confetti === 'function') confetti({ particleCount: 100, spread: 70, origin: { y: 0.5 } });
}

function continueMcSession() {
  updateMcSetupUI();
}

function finishMcSession() {
  updateMcSetupUI();
}

/* ==========================================
   TAB 2: SẮP XẾP CÂU (CUSTOM SESSION)
   ========================================== */
let builderList = [];
let builderGlobalIndex = 0;
let builderSessionList = [];
let builderSessionIndex = 0;
let builderCorrectFirstTry = 0;
let builderWrongList = JSON.parse(localStorage.getItem('hsk1_builder_wrong') || '[]');
let isBuilderWrongMode = false;
let currentCorrectWords = [];
let userSelectedPool = [];

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

  builderGlobalIndex = savedIndex ? parseInt(savedIndex, 10) : 0;
  if (builderGlobalIndex >= builderList.length) builderGlobalIndex = 0;

  updateBuilderSetupUI();
}

function updateBuilderSetupUI() {
  document.getElementById('builder-total-count').innerText = builderList.length;
  document.getElementById('builder-learned-count').innerText = builderGlobalIndex;
  document.getElementById('builder-remain-count').innerText = builderList.length - builderGlobalIndex;

  const wrongBtn = document.getElementById('builder-wrong-btn');
  if (builderWrongList.length > 0) {
    wrongBtn.style.display = 'inline-block';
    document.getElementById('builder-wrong-count').innerText = builderWrongList.length;
  } else {
    wrongBtn.style.display = 'none';
  }

  document.getElementById('builder-setup-screen').style.display = 'block';
  document.getElementById('builder-quiz-screen').style.display = 'none';
  document.getElementById('builder-summary-screen').style.display = 'none';
}

function startBuilderSession() {
  const countInput = parseInt(document.getElementById('builder-session-input').value, 10) || 10;

  if (builderGlobalIndex >= builderList.length) builderGlobalIndex = 0;

  builderSessionList = builderList.slice(builderGlobalIndex, builderGlobalIndex + countInput);
  if (builderSessionList.length === 0) {
    alert("Bạn đã ghép hết kho câu ví dụ! Hệ thống sẽ reset lại từ đầu.");
    builderGlobalIndex = 0;
    localStorage.setItem('hsk1_builder_index', 0);
    builderSessionList = builderList.slice(0, countInput);
  }

  builderSessionIndex = 0;
  builderCorrectFirstTry = 0;
  isBuilderWrongMode = false;

  document.getElementById('builder-setup-screen').style.display = 'none';
  document.getElementById('builder-quiz-screen').style.display = 'block';
  document.getElementById('builder-mode-badge').innerText = 'Phiên ghép câu';

  loadBuilderQuestion();
}

function startBuilderWrongReview() {
  if (builderWrongList.length === 0) return;

  builderSessionList = [...builderWrongList];
  builderSessionIndex = 0;
  builderCorrectFirstTry = 0;
  isBuilderWrongMode = true;

  document.getElementById('builder-setup-screen').style.display = 'none';
  document.getElementById('builder-quiz-screen').style.display = 'block';
  document.getElementById('builder-mode-badge').innerText = 'Ôn câu sai ⚠️';

  loadBuilderQuestion();
}

function loadBuilderQuestion() {
  document.getElementById('builder-feedback').innerText = '';

  if (builderSessionIndex >= builderSessionList.length) {
    showBuilderSummary();
    return;
  }

  const item = builderSessionList[builderSessionIndex];

  document.getElementById('builder-progress-text').innerText = `${builderSessionIndex + 1}/${builderSessionList.length}`;
  const pct = Math.round(((builderSessionIndex + 1) / builderSessionList.length) * 100);
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

let builderHadWrongInCurrent = false;

function checkBuilderAnswer() {
  const feedback = document.getElementById('builder-feedback');
  const userResult = userSelectedPool.map(w => w.text).join(' ');
  const correctResult = currentCorrectWords.join(' ');

  if (userSelectedPool.length === 0) {
    feedback.style.color = '#d1242f';
    feedback.innerText = '⚠️ Bạn chưa chọn từ nào kìa!';
    return;
  }

  const currentItem = builderSessionList[builderSessionIndex];

  if (userResult.toLowerCase() === correctResult.toLowerCase()) {
    feedback.style.color = '#1f883d';
    feedback.innerText = `🎉 Chuẩn đét! Đáp án: "${correctResult}"`;
    if (typeof confetti === 'function') confetti({ particleCount: 70, spread: 60, origin: { y: 0.8 } });

    if (!builderHadWrongInCurrent) builderCorrectFirstTry++;

    if (isBuilderWrongMode) {
      builderWrongList = builderWrongList.filter(item => item.exampleVn !== currentItem.exampleVn);
      localStorage.setItem('hsk1_builder_wrong', JSON.stringify(builderWrongList));
    }

    setTimeout(() => {
      builderSessionIndex++;
      if (!isBuilderWrongMode) {
        builderGlobalIndex++;
        localStorage.setItem('hsk1_builder_index', builderGlobalIndex);
      }
      builderHadWrongInCurrent = false;
      loadBuilderQuestion();
    }, 1200);

  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Sai thứ tự từ rồi, thử lại nhé!`;
    builderHadWrongInCurrent = true;

    if (!builderWrongList.some(item => item.exampleVn === currentItem.exampleVn)) {
      builderWrongList.push(currentItem);
      localStorage.setItem('hsk1_builder_wrong', JSON.stringify(builderWrongList));
    }
  }
}

function skipBuilderSentence() {
  builderSessionIndex++;
  if (!isBuilderWrongMode) {
    builderGlobalIndex++;
    localStorage.setItem('hsk1_builder_index', builderGlobalIndex);
  }
  builderHadWrongInCurrent = false;
  loadBuilderQuestion();
}

function showBuilderSummary() {
  document.getElementById('builder-quiz-screen').style.display = 'none';
  document.getElementById('builder-summary-screen').style.display = 'block';

  document.getElementById('builder-sum-done').innerText = builderSessionList.length;
  document.getElementById('builder-sum-correct').innerText = builderCorrectFirstTry;
  document.getElementById('builder-sum-wrong').innerText = builderSessionList.length - builderCorrectFirstTry;

  if (typeof confetti === 'function') confetti({ particleCount: 100, spread: 70, origin: { y: 0.5 } });
}

function continueBuilderSession() {
  updateBuilderSetupUI();
}

function finishBuilderSession() {
  updateBuilderSetupUI();
}

/* Dark Mode */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
