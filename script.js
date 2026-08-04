let hskData = [];

// Khởi tạo app
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
        initMultipleChoiceQuiz();
        initSentenceScrambleQuiz();
      }
    })
    .catch(err => console.error("Lỗi đọc file Excel:", err));
});

/* ==========================================
   DẠNG 1: TRẮC NGHIỆM CHỌN HÁN TỰ + PINYIN
   ========================================== */
let currentMcCorrectItem = null;

function initMultipleChoiceQuiz() {
  document.getElementById('mc-feedback').innerText = '';
  
  currentMcCorrectItem = hskData[Math.floor(Math.random() * hskData.length)];
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
    
    if (typeof confetti === 'function') {
      confetti({ particleCount: 80, spread: 60, origin: { y: 0.3 } });
    }

    setTimeout(initMultipleChoiceQuiz, 1500);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Sai rồi, thử lại xem!';
  }
}

/* ==========================================
   DẠNG 2: SẮP XẾP CÂU PINYIN + TIMER 10S
   ========================================== */
let currentTargetSentence = "";
let currentWords = [];
let userSelectedWords = [];
let timerInterval = null;
let timeLeft = 10; // 10 giây
let isQuizActive = true;

function initSentenceScrambleQuiz() {
  clearInterval(timerInterval); // Reset timer cũ
  isQuizActive = true;
  timeLeft = 10;
  
  document.getElementById('scramble-feedback').innerText = '';
  document.getElementById('user-sentence').innerHTML = '<span class="placeholder-text">Click chọn các từ pinyin bên dưới...</span>';
  userSelectedWords = [];

  const validItems = hskData.filter(item => item.examplePinyin && item.exampleVn);
  if (validItems.length === 0) return;

  const selectedItem = validItems[Math.floor(Math.random() * validItems.length)];
  
  document.getElementById('scramble-meaning').innerText = `"${selectedItem.exampleVn}"`;
  currentTargetSentence = selectedItem.examplePinyin.replace(/[。!？,.]/g, '').trim();

  let rawWords = currentTargetSentence.split(' ').filter(w => w.trim() !== '');
  currentWords = [...rawWords].sort(() => Math.random() - 0.5);

  renderWordPool();
  startTimer();
}

function startTimer() {
  const timerBar = document.getElementById('timer-bar');
  timerBar.style.width = '100%';
  timerBar.style.backgroundColor = '#2da44e';

  timerInterval = setInterval(() => {
    timeLeft -= 0.1;
    const percentage = (timeLeft / 10) * 100;
    timerBar.style.width = `${percentage}%`;

    // Cảnh báo màu đỏ khi còn dưới 3 giây
    if (timeLeft <= 3) {
      timerBar.style.backgroundColor = '#cf222e';
    }

    // Hết giờ!
    if (timeLeft <= 0) {
      clearInterval(timerInterval);
      isQuizActive = false;
      showTimeOutResult();
    }
  }, 100);
}

function renderWordPool() {
  const poolContainer = document.getElementById('word-pool');
  poolContainer.innerHTML = '';

  currentWords.forEach((word) => {
    const chip = document.createElement('button');
    chip.className = 'word-chip';
    chip.innerText = word;
    chip.onclick = () => selectWord(word, chip);
    poolContainer.appendChild(chip);
  });
}

function selectWord(word, btnElement) {
  if (!isQuizActive) return; // Nếu hết giờ thì khóa không cho bấm

  userSelectedWords.push(word);
  btnElement.style.visibility = 'hidden';

  const builderBox = document.getElementById('user-sentence');
  if (userSelectedWords.length === 1) {
    builderBox.innerHTML = '';
  }

  const selectedChip = document.createElement('span');
  selectedChip.className = 'word-chip';
  selectedChip.innerText = word;
  builderBox.appendChild(selectedChip);

  if (userSelectedWords.length === currentWords.length) {
    clearInterval(timerInterval); // Dừng đồng hồ khi đã xếp xong
    checkSentenceAnswer();
  }
}

function checkSentenceAnswer() {
  const userResult = userSelectedWords.join(' ');
  const feedback = document.getElementById('scramble-feedback');

  if (userResult === currentTargetSentence) {
    feedback.style.color = '#1f883d';
    feedback.innerText = '⚡ Quá đỉnh! Phản xạ xuất sắc!';

    if (typeof confetti === 'function') {
      confetti({ particleCount: 100, spread: 70, origin: { y: 0.8 } });
    }

    setTimeout(initSentenceScrambleQuiz, 2000);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Sai rồi! Đáp án đúng: ${currentTargetSentence}`;
  }
}

function showTimeOutResult() {
  const feedback = document.getElementById('scramble-feedback');
  feedback.style.color = '#d1242f';
  feedback.innerText = `⏰ Hết 10s rồi! Đáp án đúng: "${currentTargetSentence}"`;
  
  // Hiện đáp án đúng lên khung ghép câu
  const builderBox = document.getElementById('user-sentence');
  builderBox.innerHTML = `<span style="color:#0969da; font-weight:bold; font-size:18px;">${currentTargetSentence}</span>`;
}

/* ==========================================
   DARK MODE
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
