/* ==========================================
   1. TRẠNG THÁI ỨNG DỤNG (STATE MANAGEMENT)
   ========================================== */
let rawData = [];

// Từ vựng
let vocabDeck = [];
let currentVocabIndex = 0;
let vocabCorrectCount = 0;
let vocabErrorBank = JSON.parse(localStorage.getItem('vocabErrorBank')) || [];

// Ghép câu
let sentenceDeck = [];
let currentSentenceIndex = 0;
let sentenceCorrectCount = 0;
let sentenceErrorBank = JSON.parse(localStorage.getItem('sentenceErrorBank')) || [];
let userSelectedWords = [];

/* ==========================================
   2. KHỞI TẠO DỮ LIỆU TỪ EXCEL
   ========================================== */
window.addEventListener('DOMContentLoaded', () => {
  fetch('HSK1_flashcards.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      rawData = XLSX.utils.sheet_to_json(firstSheet);

      updateDashboardStats();
    })
    .catch(err => console.error("Lỗi tải file Excel:", err));
});

function updateDashboardStats() {
  const totalElem = document.getElementById('total-words');
  const learnedElem = document.getElementById('learned-words');
  const remainingElem = document.getElementById('remaining-words');

  if (totalElem) totalElem.innerText = rawData.length;
  if (learnedElem) learnedElem.innerText = 0;
  if (remainingElem) remainingElem.innerText = rawData.length;
}

/* ==========================================
   3. QUẢN LÝ TAB
   ========================================== */
function switchTab(tabName) {
  const tabs = document.querySelectorAll('.tab-content');
  const buttons = document.querySelectorAll('.tab-btn');

  tabs.forEach(tab => tab.style.display = 'none');
  buttons.forEach(btn => btn.classList.remove('active'));

  const activeTab = document.getElementById(`${tabName}-tab`);
  const activeBtn = document.getElementById(`btn-${tabName}`);

  if (activeTab) activeTab.style.display = 'block';
  if (activeBtn) activeBtn.classList.add('active');

  if (tabName === 'sentence' && sentenceDeck.length === 0) {
    startSentenceSession();
  }
}

/* ==========================================
   4. MODULE 1: VOCABULARY QUIZ (TỪ VỰNG)
   ========================================== */
function startVocabSession() {
  if (!rawData || rawData.length === 0) return;

  const inputElem = document.getElementById('vocab-session-count');
  const sessionSize = inputElem ? parseInt(inputElem.value) || 10 : 10;

  vocabDeck = [...rawData].sort(() => Math.random() - 0.5).slice(0, Math.min(sessionSize, rawData.length));
  currentVocabIndex = 0;
  vocabCorrectCount = 0;

  const setupView = document.getElementById('vocab-setup');
  const quizView = document.getElementById('vocab-quiz');

  if (setupView) setupView.style.display = 'none';
  if (quizView) quizView.style.display = 'block';

  renderVocabCard();
}

function renderVocabCard() {
  if (currentVocabIndex >= vocabDeck.length) {
    finishVocabSession();
    return;
  }

  const currentItem = vocabDeck[currentVocabIndex];

  const hanziElem = document.getElementById('vocab-hanzi');
  const pinyinElem = document.getElementById('vocab-pinyin');

  if (hanziElem) hanziElem.innerText = currentItem.Hanzi || '';
  if (pinyinElem) pinyinElem.innerText = currentItem.Pinyin || '';

  const wrongOptions = rawData
    .filter(item => item.Meaning !== currentItem.Meaning)
    .sort(() => Math.random() - 0.5)
    .slice(0, 3)
    .map(item => item.Meaning);

  const options = [...wrongOptions, currentItem.Meaning].sort(() => Math.random() - 0.5);

  const optionsContainer = document.getElementById('vocab-options');
  if (optionsContainer) {
    optionsContainer.innerHTML = '';
    options.forEach(opt => {
      const btn = document.createElement('button');
      btn.className = 'option-btn';
      btn.innerText = opt;
      btn.onclick = () => checkVocabAnswer(opt, currentItem.Meaning, btn);
      optionsContainer.appendChild(btn);
    });
  }

  updateProgress('vocab', currentVocabIndex + 1, vocabDeck.length);
}

function checkVocabAnswer(selected, correct, btnElement) {
  const buttons = document.querySelectorAll('#vocab-options .option-btn');
  buttons.forEach(b => b.disabled = true);

  if (selected === correct) {
    btnElement.classList.add('correct');
    vocabCorrectCount++;
  } else {
    btnElement.classList.add('wrong');

    const currentItem = vocabDeck[currentVocabIndex];
    if (!vocabErrorBank.some(e => e.Hanzi === currentItem.Hanzi)) {
      vocabErrorBank.push(currentItem);
      localStorage.setItem('vocabErrorBank', JSON.stringify(vocabErrorBank));
    }

    buttons.forEach(b => {
      if (b.innerText === correct) b.classList.add('correct');
    });
  }

  setTimeout(() => {
    currentVocabIndex++;
    renderVocabCard();
  }, 1200);
}

function finishVocabSession() {
  const accuracy = Math.round((vocabCorrectCount / vocabDeck.length) * 100);

  alert(`🎉 Hoàn thành phiên học!\nĐộ chính xác: ${accuracy}%`);

  const setupView = document.getElementById('vocab-setup');
  const quizView = document.getElementById('vocab-quiz');

  if (setupView) setupView.style.display = 'block';
  if (quizView) quizView.style.display = 'none';
}

/* ==========================================
   5. MODULE 2: SENTENCE BUILDER (GHẾP CÂU)
   ========================================== */
function startSentenceSession() {
  if (!rawData || rawData.length === 0) return;

  const sentenceData = rawData.filter(item => item.SentenceHanzi);
  sentenceDeck = [...sentenceData].sort(() => Math.random() - 0.5).slice(0, 5);
  currentSentenceIndex = 0;
  sentenceCorrectCount = 0;

  renderSentenceCard();
}

function renderSentenceCard() {
  if (currentSentenceIndex >= sentenceDeck.length) {
    finishSentenceSession();
    return;
  }

  userSelectedWords = [];
  const currentItem = sentenceDeck[currentSentenceIndex];

  const meaningElem = document.getElementById('sentence-meaning');
  const userZone = document.getElementById('user-sentence-zone');

  if (meaningElem) meaningElem.innerText = currentItem.SentenceMeaning || '';
  if (userZone) userZone.innerHTML = '';

  const correctWords = (currentItem.SentenceHanzi || '').split(' ');
  const shuffledWords = [...correctWords].sort(() => Math.random() - 0.5);

  const poolContainer = document.getElementById('sentence-word-pool');
  if (poolContainer) {
    poolContainer.innerHTML = '';
    shuffledWords.forEach((word, idx) => {
      const chip = document.createElement('div');
      chip.className = 'word-chip';
      chip.innerText = word;
      chip.dataset.id = idx;
      chip.onclick = () => selectWord(chip, word);
      poolContainer.appendChild(chip);
    });
  }

  updateProgress('sentence', currentSentenceIndex + 1, sentenceDeck.length);
}

function selectWord(chipElement, word) {
  if (chipElement.classList.contains('used')) return;

  chipElement.classList.add('used');
  userSelectedWords.push({ word, chipId: chipElement.dataset.id });

  const userZone = document.getElementById('user-sentence-zone');
  if (userZone) {
    const selectedChip = document.createElement('div');
    selectedChip.className = 'word-chip selected';
    selectedChip.innerText = word;
    selectedChip.onclick = () => deselectWord(selectedChip, chipElement, word);
    userZone.appendChild(selectedChip);
  }
}

function deselectWord(selectedChip, originalChip, word) {
  selectedChip.remove();
  originalChip.classList.remove('used');
  userSelectedWords = userSelectedWords.filter(item => item.chipId !== originalChip.dataset.id);
}

function checkSentenceAnswer() {
  const currentItem = sentenceDeck[currentSentenceIndex];
  const targetSentence = (currentItem.SentenceHanzi || '').replace(/\s+/g, '');
  const userSentence = userSelectedWords.map(i => i.word).join('');

  if (userSentence === targetSentence) {
    sentenceCorrectCount++;
    setTimeout(() => {
      currentSentenceIndex++;
      renderSentenceCard();
    }, 1000);
  } else {
    if (!sentenceErrorBank.some(e => e.SentenceHanzi === currentItem.SentenceHanzi)) {
      sentenceErrorBank.push(currentItem);
      localStorage.setItem('sentenceErrorBank', JSON.stringify(sentenceErrorBank));
    }

    const userZone = document.getElementById('user-sentence-zone');
    if (userZone) {
      userZone.classList.add('shake');
      setTimeout(() => userZone.classList.remove('shake'), 500);
    }
  }
}

function finishSentenceSession() {
  const accuracy = Math.round((sentenceCorrectCount / sentenceDeck.length) * 100);
  alert(`🎉 Hoàn thành ghép câu!\nĐộ chính xác: ${accuracy}%`);
  startSentenceSession();
}

/* ==========================================
   6. TIỆN ÍCH CHUNG
   ========================================== */
function updateProgress(module, current, total) {
  const percent = total > 0 ? (current / total) * 100 : 0;
  const bar = document.getElementById(`${module}-progress-bar`);
  const text = document.getElementById(`${module}-progress-text`);
  if (bar) bar.style.width = `${percent}%`;
  if (text) text.innerText = `${current}/${total}`;
}

// Khai báo hàm dùng trực tiếp từ HTML (onclick)
window.switchTab = switchTab;
window.startVocabSession = startVocabSession;
window.startSentenceSession = startSentenceSession;
window.checkSentenceAnswer = checkSentenceAnswer;
