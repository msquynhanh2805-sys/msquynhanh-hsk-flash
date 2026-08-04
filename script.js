/* ==========================================
   1. TRẠNG THÁI ỨNG DỤNG (STATE MANAGEMENT)
   ========================================== */
let rawData = [];

// Module Từ vựng
let vocabDeck = [];
let currentVocabIndex = 0;
let vocabCorrectCount = 0;
let vocabErrorBank = JSON.parse(localStorage.getItem('vocabErrorBank')) || [];

// Module Ghép câu
let sentenceDeck = [];
let currentSentenceIndex = 0;
let sentenceCorrectCount = 0;
let sentenceErrorBank = JSON.parse(localStorage.getItem('sentenceErrorBank')) || [];
let userSelectedWords = [];

/* ==========================================
   2. KHỞI TẠO DỮ LIỆU TỪ EXCEL & CHUYỂN TAB
   ========================================== */
window.addEventListener('DOMContentLoaded', () => {
  fetch('HSK1_flashcards.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      rawData = XLSX.utils.sheet_to_json(firstSheet);
      
      initVocabModule();
      initSentenceModule();
    })
    .catch(err => console.error("Lỗi tải file Excel:", err));

  setupTabListeners();
});

// Chuyển tab linh hoạt
function setupTabListeners() {
  document.addEventListener('click', (e) => {
    const target = e.target.closest('[data-tab], .tab-btn, .tab-link');
    if (!target) return;

    const tabAttr = target.getAttribute('data-tab') || target.id;
    if (tabAttr) {
      switchTab(tabAttr);
    }
  });
}

function switchTab(tabIdentifier) {
  const isSentence = tabIdentifier.toLowerCase().includes('sentence') || tabIdentifier.toLowerCase().includes('builder');

  const vocabTab = document.getElementById('vocab-tab') || document.getElementById('vocab-module') || document.querySelector('.vocab-section');
  const sentenceTab = document.getElementById('sentence-tab') || document.getElementById('sentence-module') || document.querySelector('.sentence-section');

  if (vocabTab && sentenceTab) {
    if (isSentence) {
      vocabTab.style.display = 'none';
      sentenceTab.style.display = 'block';
    } else {
      vocabTab.style.display = 'block';
      sentenceTab.style.display = 'none';
    }
  }

  document.querySelectorAll('.tab-btn, .tab-link').forEach(btn => btn.classList.remove('active'));
  const activeBtn = document.querySelector(`[data-tab*="${isSentence ? 'sentence' : 'vocab'}"]`) || document.getElementById(tabIdentifier);
  if (activeBtn) activeBtn.classList.add('active');
}

/* ==========================================
   3. MODULE 1: VOCABULARY QUIZ (TỪ VỰNG)
   ========================================== */
function initVocabModule() {
  if (!rawData || rawData.length === 0) return;
  vocabDeck = [...rawData].sort(() => Math.random() - 0.5).slice(0, 10);
  currentVocabIndex = 0;
  vocabCorrectCount = 0;
  renderVocabCard();
}

function renderVocabCard() {
  if (currentVocabIndex >= vocabDeck.length) {
    finishVocabSession();
    return;
  }

  const currentItem = vocabDeck[currentVocabIndex];
  
  const hanziElem = document.getElementById('vocab-hanzi') || document.querySelector('.hanzi-display');
  const pinyinElem = document.getElementById('vocab-pinyin') || document.querySelector('.pinyin-display');
  
  if (hanziElem) hanziElem.innerText = currentItem.Hanzi || '';
  if (pinyinElem) pinyinElem.innerText = currentItem.Pinyin || '';
  
  const wrongOptions = rawData
    .filter(item => item.Meaning !== currentItem.Meaning)
    .sort(() => Math.random() - 0.5)
    .slice(0, 3)
    .map(item => item.Meaning);
    
  const options = [...wrongOptions, currentItem.Meaning].sort(() => Math.random() - 0.5);
  
  const optionsContainer = document.getElementById('vocab-options') || document.querySelector('.options-grid');
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
  const buttons = document.querySelectorAll('.option-btn');
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
  }, 1000);
}

function finishVocabSession() {
  const accuracy = Math.round((vocabCorrectCount / vocabDeck.length) * 100);
  showFeedbackModal('vocab', accuracy);
}

/* ==========================================
   4. MODULE 2: SENTENCE BUILDER (GHIÉP CÂU)
   ========================================== */
function initSentenceModule() {
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
  
  const meaningElem = document.getElementById('sentence-meaning') || document.querySelector('.sentence-meaning-display');
  const userZone = document.getElementById('user-sentence-zone') || document.querySelector('.answer-zone');
  
  if (meaningElem) meaningElem.innerText = currentItem.SentenceMeaning || '';
  if (userZone) userZone.innerHTML = '';
  
  const correctWords = (currentItem.SentenceHanzi || '').split(' ');
  const shuffledWords = [...correctWords].sort(() => Math.random() - 0.5);
  
  const poolContainer = document.getElementById('sentence-word-pool') || document.querySelector('.word-pool');
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
  
  const userZone = document.getElementById('user-sentence-zone') || document.querySelector('.answer-zone');
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
    }, 800);
  } else {
    if (!sentenceErrorBank.some(e => e.SentenceHanzi === currentItem.SentenceHanzi)) {
      sentenceErrorBank.push(currentItem);
      localStorage.setItem('sentenceErrorBank', JSON.stringify(sentenceErrorBank));
    }
    
    const userZone = document.getElementById('user-sentence-zone') || document.querySelector('.answer-zone');
    if (userZone) {
      userZone.classList.add('shake');
      setTimeout(() => userZone.classList.remove('shake'), 500);
    }
  }
}

function finishSentenceSession() {
  const accuracy = Math.round((sentenceCorrectCount / sentenceDeck.length) * 100);
  showFeedbackModal('sentence', accuracy);
}

/* ==========================================
   5. TIỆN ÍCH CHUNG
   ========================================== */
function updateProgress(module, current, total) {
  const percent = total > 0 ? (current / total) * 100 : 0;
  const bar = document.getElementById(`${module}-progress-bar`) || document.querySelector(`.${module}-progress-bar`);
  const text = document.getElementById(`${module}-progress-text`) || document.querySelector(`.${module}-progress-text`);
  if (bar) bar.style.width = `${percent}%`;
  if (text) text.innerText = `${current}/${total}`;
}

function showFeedbackModal(type, accuracy) {
  let title = '';
  let message = '';

  if (accuracy >= 80) {
    title = '🎉 Xuất Sắc!';
    message = `Đạt ${accuracy}% độ chính xác.`;
  } else if (accuracy >= 50) {
    title = '👍 Hoàn Thành!';
    message = `Đạt ${accuracy}%.`;
  } else {
    title = '💪 Cố Gắng Lên!';
    message = `Đạt ${accuracy}%. Cùng ôn lại nhé!`;
  }

  alert(`${title}\n${message}`);
  
  if (type === 'vocab') initVocabModule();
  else initSentenceModule();
}

// Khai báo global
window.switchTab = switchTab;
window.startVocabSession = initVocabModule;
window.startSentenceSession = initSentenceModule;
window.checkSentenceAnswer = checkSentenceAnswer;
