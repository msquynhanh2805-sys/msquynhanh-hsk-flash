/* ==========================================
   1. BỘ TẠO ÂM THANH PIANO (ACOUSTIC PIANO)
   ========================================== */
const AudioFX = {
  ctx: null,
  init() {
    if (!this.ctx) {
      this.ctx = new (window.AudioContext || window.webkitAudioContext)();
    }
    if (this.ctx.state === 'suspended') {
      this.ctx.resume();
    }
  },

  playPianoNote(freq, duration = 0.8, volume = 0.15) {
    try {
      this.init();
      const now = this.ctx.currentTime;

      const osc1 = this.ctx.createOscillator();
      const gain1 = this.ctx.createGain();
      osc1.type = 'triangle';
      osc1.frequency.setValueAtTime(freq, now);

      const osc2 = this.ctx.createOscillator();
      const gain2 = this.ctx.createGain();
      osc2.type = 'sine';
      osc2.frequency.setValueAtTime(freq * 2, now);

      const hammer = this.ctx.createOscillator();
      const hammerGain = this.ctx.createGain();
      hammer.type = 'sine';
      hammer.frequency.setValueAtTime(120, now);

      gain1.gain.setValueAtTime(volume, now);
      gain1.gain.exponentialRampToValueAtTime(0.0001, now + duration);

      gain2.gain.setValueAtTime(volume * 0.3, now);
      gain2.gain.exponentialRampToValueAtTime(0.0001, now + (duration * 0.7));

      hammerGain.gain.setValueAtTime(volume * 0.2, now);
      hammerGain.gain.exponentialRampToValueAtTime(0.0001, now + 0.03);

      osc1.connect(gain1);
      osc2.connect(gain2);
      hammer.connect(hammerGain);

      gain1.connect(this.ctx.destination);
      gain2.connect(this.ctx.destination);
      hammerGain.connect(this.ctx.destination);

      osc1.start(now);
      osc2.start(now);
      hammer.start(now);

      osc1.stop(now + duration);
      osc2.stop(now + duration);
      hammer.stop(now + 0.03);
    } catch (e) {
      console.log("Audio Error:", e);
    }
  },

  playCorrect() {
    this.playPianoNote(523.25, 0.7, 0.18); 
  },

  playWrong() {
    this.playPianoNote(174.61, 0.4, 0.12);
  },

  playCelebration() {
    const notes = [523.25, 659.25, 783.99, 1046.50];
    notes.forEach((freq, idx) => {
      setTimeout(() => {
        this.playPianoNote(freq, 0.9, 0.12);
      }, idx * 120);
    });
  }
};

/* ==========================================
   2. TRẠNG THÁI DỮ LIỆU (STATE MANAGEMENT)
   ========================================== */
let rawData = [];
let vocabDeck = [];
let currentVocabIndex = 0;
let vocabCorrectCount = 0;
let vocabErrorBank = JSON.parse(localStorage.getItem('vocabErrorBank')) || [];

let sentenceDeck = [];
let currentSentenceIndex = 0;
let sentenceCorrectCount = 0;
let sentenceErrorBank = JSON.parse(localStorage.getItem('sentenceErrorBank')) || [];
let userSelectedWords = [];

/* ==========================================
   3. KHỞI TẠO TỰ ĐỘNG & BẮT SỰ KIỆN CLICK
   ========================================== */
window.addEventListener('DOMContentLoaded', () => {
  // 1. Đọc file Excel
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

  // 2. Tự động nhận diện các Tab nút bấm trên trang
  setupTabListeners();
});

// Tự động tìm và gán sự kiện cho các Tab
function setupTabListeners() {
  document.addEventListener('click', (e) => {
    AudioFX.init(); // Đảm bảo bật Audio khi có tương tác

    // Tìm xem người dùng có bấm vào nút tab hay không
    const target = e.target.closest('[data-tab], .tab-btn, .tab-link, button');
    if (!target) return;

    // Chuyển Tab dựa theo thuộc tính data-tab hoặc Text của nút
    const tabAttr = target.getAttribute('data-tab') || target.id;
    
    if (tabAttr && (tabAttr.includes('vocab') || tabAttr.includes('sentence') || tabAttr.includes('quiz') || tabAttr.includes('builder'))) {
      switchTab(tabAttr);
    }
  });
}

// Hàm chuyển Tab linh hoạt
function switchTab(tabIdentifier) {
  const isSentence = tabIdentifier.toLowerCase().includes('sentence') || tabIdentifier.toLowerCase().includes('builder');
  
  // Ẩn tất cả các tab
  const allTabs = document.querySelectorAll('.tab-content, section, .module-container');
  allTabs.forEach(tab => {
    if (tab.id) tab.style.display = 'none';
  });

  // Hiện tab tương ứng
  const targetVocab = document.getElementById('vocab-tab') || document.getElementById('vocab-module') || document.querySelector('.vocab-section');
  const targetSentence = document.getElementById('sentence-tab') || document.getElementById('sentence-module') || document.querySelector('.sentence-section');

  if (isSentence && targetSentence) {
    targetSentence.style.display = 'block';
  } else if (!isSentence && targetVocab) {
    targetVocab.style.display = 'block';
  }

  // Active Style cho nút bấm
  document.querySelectorAll('.tab-btn, .tab-link').forEach(btn => btn.classList.remove('active'));
  const activeBtn = document.querySelector(`[data-tab*="${isSentence ? 'sentence' : 'vocab'}"]`) || document.getElementById(tabIdentifier);
  if (activeBtn) activeBtn.classList.add('active');
}

/* ==========================================
   4. MODULE 1: VOCABULARY QUIZ
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
  AudioFX.init();
  const buttons = document.querySelectorAll('.option-btn');
  buttons.forEach(b => b.disabled = true);

  if (selected === correct) {
    btnElement.classList.add('correct');
    AudioFX.playCorrect();
    vocabCorrectCount++;
  } else {
    btnElement.classList.add('wrong');
    AudioFX.playWrong();
    
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
  showFeedbackModal('vocab', accuracy);
}

/* ==========================================
   5. MODULE 2: SENTENCE BUILDER
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
  AudioFX.init();
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
  AudioFX.init();
  const currentItem = sentenceDeck[currentSentenceIndex];
  const targetSentence = (currentItem.SentenceHanzi || '').replace(/\s+/g, '');
  const userSentence = userSelectedWords.map(i => i.word).join('');

  if (userSentence === targetSentence) {
    AudioFX.playCorrect();
    sentenceCorrectCount++;
    setTimeout(() => {
      currentSentenceIndex++;
      renderSentenceCard();
    }, 1000);
  } else {
    AudioFX.playWrong();
    
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
   6. TIỆN ÍCH CHUNG (PROGRESS & MODAL)
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
    title = '🎉 Cực Kỳ Xuất Sắc!';
    message = `Bạn đạt ${accuracy}% độ chính xác. Tiếp tục duy trì phong độ nhé!`;
    AudioFX.playCelebration();
    if (typeof confetti === 'function') confetti();
  } else if (accuracy >= 50) {
    title = '👍 Làm Tốt Lắm!';
    message = `Bạn đạt ${accuracy}%. Ôn lại các câu sai để hoàn thiện hơn nhé.`;
    AudioFX.playCorrect();
  } else {
    title = '💪 Cố Gắng Lên!';
    message = `Đạt ${accuracy}%. Đừng nản lòng, luyện tập thêm phiên nữa nhé!`;
    AudioFX.playWrong();
  }

  alert(`${title}\n${message}`);
  
  if (type === 'vocab') initVocabModule();
  else initSentenceModule();
}

// Khai báo các hàm toàn cục để tương thích với mọi nút onclick trong HTML
window.switchTab = switchTab;
window.startVocabSession = initVocabModule;
window.startSentenceSession = initSentenceModule;
window.checkSentenceAnswer = checkSentenceAnswer;
