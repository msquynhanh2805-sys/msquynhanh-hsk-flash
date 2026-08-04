/* ==========================================
   1. BỘ TẠO ÂM THANH PIANO (ACOUSTIC PIANO SYNTH)
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
      console.log("Lỗi âm thanh:", e);
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
   2. TRẠNG THÁI ỨNG DỤNG (STATE MANAGEMENT)
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
   3. KHỞI TẠO DỮ LIỆU & SỰ KIỆN NÚT BẤM
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
});

// Hàm hỗ trợ các nút bấm bắt đầu / chọn lại
function startVocabSession() {
  AudioFX.init();
  initVocabModule();
}

function startSentenceSession() {
  AudioFX.init();
  initSentenceModule();
}

// Chuyển đổi giữa các Tab
function switchTab(tabName) {
  AudioFX.init();
  const tabs = document.querySelectorAll('.tab-content');
  const buttons = document.querySelectorAll('.tab-btn');
  
  tabs.forEach(tab => tab.classList.remove('active'));
  buttons.forEach(btn => btn.classList.remove('active'));

  const activeTab = document.getElementById(`${tabName}-tab`);
  const activeBtn = document.getElementById(`btn-${tabName}`);
  
  if (activeTab) activeTab.classList.add('active');
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
  AudioFX.init();
  const buttons = document.querySelectorAll('#vocab-options .option-btn');
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
  
  const meaningElem = document.getElementById('sentence-meaning');
  const userZone = document.getElementById('user-sentence-zone');
  
  if (meaningElem) meaningElem.innerText = currentItem.SentenceMeaning || '';
  if (userZone) userZone.innerHTML = '';
  
  const correctWords = currentItem.SentenceHanzi.split(' ');
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
  AudioFX.init();
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
  AudioFX.init();
  const currentItem = sentenceDeck[currentSentenceIndex];
  const targetSentence = currentItem.SentenceHanzi.replace(/\s+/g, '');
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
    
    const userZone = document.getElementById('user-sentence-zone');
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
  const bar = document.getElementById(`${module}-progress-bar`);
  const text = document.getElementById(`${module}-progress-text`);
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
