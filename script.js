/* ==========================================
   1. BỘ TẠO ÂM THANH NHẸ NHÀNG (SOFT AUDIO)
   ========================================== */
const AudioFX = {
  ctx: null,
  init() {
    if (!this.ctx) {
      this.ctx = new (window.AudioContext || window.webkitAudioContext)();
    }
  },

  // Âm thanh ĐÚNG: Chuông gỗ Marimba nhẹ nhàng (C6 -> E6)
  playCorrect() {
    this.init();
    const now = this.ctx.currentTime;
    const notes = [1046.50, 1318.51];

    notes.forEach((freq, idx) => {
      const osc = this.ctx.createOscillator();
      const gain = this.ctx.createGain();

      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, now + idx * 0.08);

      gain.gain.setValueAtTime(0.12, now + idx * 0.08);
      gain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.08 + 0.3);

      osc.connect(gain);
      gain.connect(this.ctx.destination);

      osc.start(now + idx * 0.08);
      osc.stop(now + idx * 0.08 + 0.3);
    });
  },

  // Âm thanh SAI: Tiếng Pop trầm nhẹ, êm ái
  playWrong() {
    this.init();
    const now = this.ctx.currentTime;
    const osc = this.ctx.createOscillator();
    const gain = this.ctx.createGain();

    osc.type = 'sine';
    osc.frequency.setValueAtTime(220, now);
    osc.frequency.exponentialRampToValueAtTime(130, now + 0.12);

    gain.gain.setValueAtTime(0.15, now);
    gain.gain.exponentialRampToValueAtTime(0.001, now + 0.12);

    osc.connect(gain);
    gain.connect(this.ctx.destination);

    osc.start(now);
    osc.stop(now + 0.12);
  },

  // Âm thanh PHÁO HOA: Chuông gió ngân dài (Acoustic Chimes)
  playCelebration() {
    this.init();
    const now = this.ctx.currentTime;
    const chimeNotes = [523.25, 659.25, 783.99, 987.77, 1046.50];

    chimeNotes.forEach((freq, idx) => {
      const osc = this.ctx.createOscillator();
      const gain = this.ctx.createGain();

      osc.type = 'sine';
      osc.frequency.setValueAtTime(freq, now + idx * 0.1);

      gain.gain.setValueAtTime(0.1, now + idx * 0.1);
      gain.gain.exponentialRampToValueAtTime(0.001, now + idx * 0.1 + 0.6);

      osc.connect(gain);
      gain.connect(this.ctx.destination);

      osc.start(now + idx * 0.1);
      osc.stop(now + idx * 0.1 + 0.6);
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
   3. KHỞI TẠO DỮ LIỆU TỪ EXCEL
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

/* ==========================================
   4. MODULE 1: VOCABULARY QUIZ
   ========================================== */
function initVocabModule() {
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
  document.getElementById('vocab-hanzi').innerText = currentItem.Hanzi || '';
  document.getElementById('vocab-pinyin').innerText = currentItem.Pinyin || '';
  
  // Tạo 4 lựa chọn (1 đúng, 3 sai)
  const wrongOptions = rawData
    .filter(item => item.Meaning !== currentItem.Meaning)
    .sort(() => Math.random() - 0.5)
    .slice(0, 3)
    .map(item => item.Meaning);
    
  const options = [...wrongOptions, currentItem.Meaning].sort(() => Math.random() - 0.5);
  
  const optionsContainer = document.getElementById('vocab-options');
  optionsContainer.innerHTML = '';
  
  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'option-btn';
    btn.innerText = opt;
    btn.onclick = () => checkVocabAnswer(opt, currentItem.Meaning, btn);
    optionsContainer.appendChild(btn);
  });

  updateProgress('vocab', currentVocabIndex + 1, vocabDeck.length);
}

function checkVocabAnswer(selected, correct, btnElement) {
  const buttons = document.querySelectorAll('#vocab-options .option-btn');
  buttons.forEach(b => b.disabled = true);

  if (selected === correct) {
    btnElement.classList.add('correct');
    AudioFX.playCorrect();
    vocabCorrectCount++;
  } else {
    btnElement.classList.add('wrong');
    AudioFX.playWrong();
    
    // Thêm vào ngân hàng lỗi
    const currentItem = vocabDeck[currentVocabIndex];
    if (!vocabErrorBank.some(e => e.Hanzi === currentItem.Hanzi)) {
      vocabErrorBank.push(currentItem);
      localStorage.setItem('vocabErrorBank', JSON.stringify(vocabErrorBank));
    }

    // Hiển thị đáp án đúng
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
  // Lọc các hàng có chứa câu (Sentence)
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
  
  document.getElementById('sentence-meaning').innerText = currentItem.SentenceMeaning || '';
  document.getElementById('user-sentence-zone').innerHTML = '';
  
  // Tách từ/ký tự để ghép
  const correctWords = currentItem.SentenceHanzi.split(' ');
  const shuffledWords = [...correctWords].sort(() => Math.random() - 0.5);
  
  const poolContainer = document.getElementById('sentence-word-pool');
  poolContainer.innerHTML = '';
  
  shuffledWords.forEach((word, idx) => {
    const chip = document.createElement('div');
    chip.className = 'word-chip';
    chip.innerText = word;
    chip.dataset.id = idx;
    chip.onclick = () => selectWord(chip, word);
    poolContainer.appendChild(chip);
  });

  updateProgress('sentence', currentSentenceIndex + 1, sentenceDeck.length);
}

function selectWord(chipElement, word) {
  if (chipElement.classList.contains('used')) return;
  
  chipElement.classList.add('used');
  userSelectedWords.push({ word, chipId: chipElement.dataset.id });
  
  const userZone = document.getElementById('user-sentence-zone');
  const selectedChip = document.createElement('div');
  selectedChip.className = 'word-chip selected';
  selectedChip.innerText = word;
  selectedChip.onclick = () => deselectWord(selectedChip, chipElement, word);
  userZone.appendChild(selectedChip);
}

function deselectWord(selectedChip, originalChip, word) {
  selectedChip.remove();
  originalChip.classList.remove('used');
  userSelectedWords = userSelectedWords.filter(item => item.chipId !== originalChip.dataset.id);
}

function checkSentenceAnswer() {
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
    userZone.classList.add('shake');
    setTimeout(() => userZone.classList.remove('shake'), 500);
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
  const percent = (current / total) * 100;
  document.getElementById(`${module}-progress-bar`).style.width = `${percent}%`;
  document.getElementById(`${module}-progress-text`).innerText = `${current}/${total}`;
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
  
  // Reset lại module sau khi báo kết quả
  if (type === 'vocab') initVocabModule();
  else initSentenceModule();
}
