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
        initMultipleChoiceQuiz();
        initSentenceScrambleApp();
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
    if (typeof confetti === 'function') confetti({ particleCount: 80, spread: 60, origin: { y: 0.3 } });
    setTimeout(initMultipleChoiceQuiz, 1200);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Sai rồi, thử lại xem!';
  }
}

/* ==========================================
   DẠNG 2: SẮP XẾP CÂU (LƯU TIẾN ĐỘ + 60s + PAUSE + HOÀN TỪ)
   ========================================== */
let sentenceList = [];
let currentSentenceIndex = 0;

let currentTargetSentence = "";
let poolWords = [];         // Các từ hiện có ở khung dưới [{id, word}]
let selectedWords = [];     // Các từ đã chọn ở khung trên [{id, word}]

let timerInterval = null;
let timeLeft = 60;          // 60 giây
let isPaused = false;
let isQuizActive = true;

function initSentenceScrambleApp() {
  // Lấy câu ví dụ có đủ Pinyin & Nghĩa VN
  const validItems = hskData.filter(item => item.examplePinyin && item.exampleVn);

  // Lấy thứ tự câu đã lưu từ LocalStorage (hoặc trộn ngẫu nhiên lần đầu)
  const savedOrder = localStorage.getItem('hsk1_sentence_order');
  const savedIndex = localStorage.getItem('hsk1_sentence_index');

  if (savedOrder) {
    const indices = JSON.parse(savedOrder);
    sentenceList = indices.map(idx => validItems[idx]).filter(Boolean);
  } else {
    // Trộn ngẫu nhiên câu lần đầu tiên
    const indices = validItems.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_sentence_order', JSON.stringify(indices));
    sentenceList = indices.map(idx => validItems[idx]);
  }

  if (savedIndex) {
    currentSentenceIndex = parseInt(savedIndex, 10);
    if (currentSentenceIndex >= sentenceList.length) {
      currentSentenceIndex = 0; // Đã xong hết thì quay lại từ đầu
    }
  }

  loadSentenceQuestion();
}

function loadSentenceQuestion() {
  clearInterval(timerInterval);
  isPaused = false;
  isQuizActive = true;
  timeLeft = 60;

  document.getElementById('btn-pause').innerText = '⏸️ Tạm dừng';
  document.getElementById('scramble-feedback').innerText = '';
  document.getElementById('scramble-progress').innerText = `Tiến độ câu: ${currentSentenceIndex + 1} / ${sentenceList.length}`;

  const item = sentenceList[currentSentenceIndex];
  document.getElementById('scramble-meaning').innerText = `"${item.exampleVn}"`;
  currentTargetSentence = item.examplePinyin.replace(/[。!？,.]/g, '').trim();

  // Tách câu thành danh sách từ và đánh ID từng từ
  const rawWords = currentTargetSentence.split(' ').filter(w => w.trim() !== '');
  const shuffled = [...rawWords].sort(() => Math.random() - 0.5);

  poolWords = shuffled.map((word, index) => ({ id: `word-${index}-${Date.now()}`, word }));
  selectedWords = [];

  renderSentenceBoxes();
  startTimer();
}

function renderSentenceBoxes() {
  // 1. Render khung ghép câu bên trên
  const builderBox = document.getElementById('user-sentence');
  builderBox.innerHTML = '';

  if (selectedWords.length === 0) {
    builderBox.innerHTML = '<span class="placeholder-text">Click chọn các từ pinyin bên dưới...</span>';
  } else {
    selectedWords.forEach((item) => {
      const chip = document.createElement('button');
      chip.className = 'word-chip selected-chip';
      chip.innerText = item.word;
      // Click vào từ ở ô câu ➔ Trả từ về ô bên dưới
      chip.onclick = () => deselectWord(item);
      builderBox.appendChild(chip);
    });
  }

  // 2. Render khung chứa từ bên dưới
  const poolContainer = document.getElementById('word-pool');
  poolContainer.innerHTML = '';

  poolWords.forEach((item) => {
    const chip = document.createElement('button');
    chip.className = 'word-chip';
    chip.innerText = item.word;
    chip.onclick = () => selectWord(item);
    poolContainer.appendChild(chip);
  });
}

function selectWord(item) {
  if (!isQuizActive || isPaused) return;

  // Chuyển từ pool -> selected
  poolWords = poolWords.filter(w => w.id !== item.id);
  selectedWords.push(item);

  renderSentenceBoxes();

  // Khi đã chọn hết từ
  if (poolWords.length === 0) {
    checkSentenceAnswer();
  }
}

function deselectWord(item) {
  if (!isQuizActive || isPaused) return;

  // Chuyển từ selected -> pool
  selectedWords = selectedWords.filter(w => w.id !== item.id);
  poolWords.push(item);

  renderSentenceBoxes();
}

function resetCurrentSentence() {
  if (!isQuizActive) return;
  
  // Trả toàn bộ từ về lại pool
  poolWords = [...poolWords, ...selectedWords];
  selectedWords = [];
  document.getElementById('scramble-feedback').innerText = '';
  renderSentenceBoxes();
}

function checkSentenceAnswer() {
  const userResult = selectedWords.map(w => w.word).join(' ');
  const feedback = document.getElementById('scramble-feedback');

  if (userResult === currentTargetSentence) {
    clearInterval(timerInterval);
    isQuizActive = false;

    feedback.style.color = '#1f883d';
    feedback.innerText = '⚡ Quá chuẩn đét! Chuẩn bị sang câu tiếp theo...';

    if (typeof confetti === 'function') confetti({ particleCount: 100, spread: 70, origin: { y: 0.8 } });

    // Lưu tiến độ câu tiếp theo
    currentSentenceIndex++;
    localStorage.setItem('hsk1_sentence_index', currentSentenceIndex);

    setTimeout(loadSentenceQuestion, 2000);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Thứ tự chưa chính xác! Click vào từ ở ô trên để bỏ bớt và chọn lại.';
  }
}

/* ==========================================
   TIMER & ĐIỀU KHIỂN
   ========================================== */
function startTimer() {
  const timerBar = document.getElementById('timer-bar');
  const timerText = document.getElementById('timer-text');

  timerInterval = setInterval(() => {
    if (isPaused || !isQuizActive) return;

    timeLeft -= 0.1;
    const percentage = (timeLeft / 60) * 100;
    
    timerBar.style.width = `${percentage}%`;
    timerText.innerText = `⏱️ ${Math.ceil(timeLeft)}s`;

    if (timeLeft <= 10) {
      timerBar.style.backgroundColor = '#cf222e';
    } else {
      timerBar.style.backgroundColor = '#2da44e';
    }

    if (timeLeft <= 0) {
      clearInterval(timerInterval);
      isQuizActive = false;
      showTimeOutResult();
    }
  }, 100);
}

function togglePause() {
  if (!isQuizActive) return;

  isPaused = !isPaused;
  const btnPause = document.getElementById('btn-pause');
  const feedback = document.getElementById('scramble-feedback');

  if (isPaused) {
    btnPause.innerText = '▶️ Tiếp tục';
    feedback.style.color = '#9a6700';
    feedback.innerText = '⏸️ Đã tạm dừng đồng hồ.';
  } else {
    btnPause.innerText = '⏸️ Tạm dừng';
    feedback.innerText = '';
  }
}

function skipToNextSentence() {
  currentSentenceIndex++;
  localStorage.setItem('hsk1_sentence_index', currentSentenceIndex);
  loadSentenceQuestion();
}

function showTimeOutResult() {
  const feedback = document.getElementById('scramble-feedback');
  feedback.style.color = '#d1242f';
  feedback.innerText = `⏰ Hết 60s rồi! Đáp án chuẩn: "${currentTargetSentence}"`;

  const builderBox = document.getElementById('user-sentence');
  builderBox.innerHTML = `<span style="color:#0969da; font-weight:bold; font-size:18px;">${currentTargetSentence}</span>`;
  document.getElementById('word-pool').innerHTML = '';
}

/* ==========================================
   DARK MODE
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
