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

/* ==========================================
   CHUYỂN TAB (TAB NAVIGATION)
   ========================================== */
function switchTab(tabId, btnElement) {
  // Ẩn tất cả tab nội dung
  document.querySelectorAll('.tab-content').forEach(tab => {
    tab.classList.remove('active');
  });

  // Bỏ trạng thái active trên nút cũ
  document.querySelectorAll('.tab-btn').forEach(btn => {
    btn.classList.remove('active');
  });

  // Mở tab được chọn
  document.getElementById(tabId).classList.add('active');
  btnElement.classList.add('active');
}

/* ==========================================
   DẠNG 1: TRẮC NGHIỆM TỪ VỰNG (LƯU TIẾN ĐỘ)
   ========================================== */
let mcList = [];
let mcIndex = 0;
let currentMcCorrectItem = null;

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

  loadMcQuestion();
}

function loadMcQuestion() {
  document.getElementById('mc-feedback').innerText = '';
  document.getElementById('mc-progress').innerText = `Tiến độ từ: ${mcIndex + 1} / ${mcList.length}`;

  currentMcCorrectItem = mcList[mcIndex];
  document.getElementById('mc-meaning').innerText = `"${currentMcCorrectItem.meaning}"`;

  const wrongOptions = hskData
    .filter(item => item.hanzi !== currentMcCorrectItem.hanzi)
    .sort(() => Math.random() - 0.5)
    .slice(0, 2);

  const options = [currentMcCorrectItem, ...wrongOptions].sort(() => Math.random() - 0.5);

  const optionsContainer = document.getElementById('mc-options');
  optionsContainer.innerHTML = '';

  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'option-btn';
    btn.onclick = () => checkMultipleChoice(opt);

    btn.innerHTML = `
      <span class="option-hanzi">${opt.hanzi}</span>
      <span class="option-pinyin">${opt.pinyin}</span>
    `;
    optionsContainer.appendChild(btn);
  });
}

function checkMultipleChoice(selectedItem) {
  const feedback = document.getElementById('mc-feedback');

  if (selectedItem.hanzi === currentMcCorrectItem.hanzi) {
    feedback.style.color = '#1f883d';
    feedback.innerText = '🎉 Chính xác!';
    if (typeof confetti === 'function') confetti({ particleCount: 80, spread: 60, origin: { y: 0.3 } });

    mcIndex++;
    localStorage.setItem('hsk1_mc_index', mcIndex);
    setTimeout(loadMcQuestion, 1200);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Sai rồi, thử lại xem!';
  }
}

function skipMcQuestion() {
  mcIndex++;
  localStorage.setItem('hsk1_mc_index', mcIndex);
  loadMcQuestion();
}

/* ==========================================
   DẠNG 2: SẮP XẾP TỪ THÀNH CÂU (SENTENCE BUILDER)
   ========================================== */
let builderList = [];
let currentBuilderIndex = 0;
let currentSentenceWords = [];
let userSelectedWords = [];

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
    currentBuilderIndex = parseInt(savedIndex, 10);
    if (currentBuilderIndex >= builderList.length) currentBuilderIndex = 0;
  }

  loadBuilderQuestion();
}

function loadBuilderQuestion() {
  document.getElementById('builder-feedback').innerText = '';
  document.getElementById('builder-progress').innerText = `Tiến độ câu: ${currentBuilderIndex + 1} / ${builderList.length}`;

  const item = builderList[currentBuilderIndex];
  document.getElementById('builder-meaning').innerText = `"${item.exampleVn}"`;

  let rawPinyin = item.examplePinyin.replace(/[。!？,.]/g, '').trim();
  currentSentenceWords = rawPinyin.split(/\s+/);

  let shuffledWords = [...currentSentenceWords].sort(() => Math.random() - 0.5);

  userSelectedWords = [];
  renderBuilderUI(shuffledWords);
}

function renderBuilderUI(poolWords) {
  const box = document.getElementById('selected-words-box');
  const pool = document.getElementById('word-pool');

  box.innerHTML = '';
  if (userSelectedWords.length === 0) {
    box.innerHTML = `<p class="placeholder-text" id="builder-placeholder">Click vào các từ bên dưới để ghép câu...</p>`;
  } else {
    userSelectedWords.forEach((wordObj, idx) => {
      const chip = document.createElement('span');
      chip.className = 'word-chip selected';
      chip.innerText = wordObj.text;
      chip.onclick = () => unselectWord(idx);
      box.appendChild(chip);
    });
  }

  pool.innerHTML = '';
  poolWords.forEach((word, idx) => {
    const chip = document.createElement('button');
    chip.className = 'word-chip';
    chip.innerText = word;
    chip.onclick = () => selectWord(word, idx, poolWords);
    pool.appendChild(chip);
  });
}

function selectWord(word, idx, currentPool) {
  userSelectedWords.push({ text: word });
  const newPool = currentPool.filter((_, i) => i !== idx);
  renderBuilderUI(newPool);
}

function unselectWord(selectedIdx) {
  const removed = userSelectedWords.splice(selectedIdx, 1)[0];
  
  const poolChips = Array.from(document.querySelectorAll('#word-pool .word-chip')).map(c => c.innerText);
  poolChips.push(removed.text);

  renderBuilderUI(poolChips);
}

function resetBuilderSelection() {
  loadBuilderQuestion();
}

function checkBuilderAnswer() {
  const feedback = document.getElementById('builder-feedback');
  const userResult = userSelectedWords.map(w => w.text).join(' ');
  const correctResult = currentSentenceWords.join(' ');

  if (userSelectedWords.length === 0) {
    feedback.style.color = '#d1242f';
    feedback.innerText = '⚠️ M chưa chọn từ nào kìa!';
    return;
  }

  if (userResult.toLowerCase() === correctResult.toLowerCase()) {
    feedback.style.color = '#1f883d';
    feedback.innerText = `🎉 Chuẩn đét! Đáp án: ${correctResult}`;
    if (typeof confetti === 'function') confetti({ particleCount: 90, spread: 70, origin: { y: 0.8 } });

    currentBuilderIndex++;
    localStorage.setItem('hsk1_builder_index', currentBuilderIndex);
    setTimeout(loadBuilderQuestion, 1800);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Chưa đúng ngữ pháp rồi, thử lại xem!`;
  }
}

function speakSample() {
  const item = builderList[currentBuilderIndex];
  if (!item || !item.examplePinyin) return;

  const textToSpeak = encodeURIComponent(item.examplePinyin);
  const audioUrl = `https://translate.google.com/translate_tts?ie=UTF-8&q=${textToSpeak}&tl=zh-CN&client=tw-ob`;

  const audio = new Audio(audioUrl);
  audio.play().catch(() => {
    const utterance = new SpeechSynthesisUtterance(item.examplePinyin);
    utterance.lang = 'zh-CN';
    window.speechSynthesis.speak(utterance);
  });
}

function skipBuilderSentence() {
  currentBuilderIndex++;
  localStorage.setItem('hsk1_builder_index', currentBuilderIndex);
  loadBuilderQuestion();
}

/* ==========================================
   DARK MODE
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
