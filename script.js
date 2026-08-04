let hskData = [];
let currentIndex = 0;

let currentTargetSentence = "";
let currentWords = [];
let userSelectedWords = [];

// Khởi tạo app
window.addEventListener('DOMContentLoaded', () => {
  fetch('HSK1_flashcards.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Lấy từ dòng thứ 2 (5 cột: Hanzi, Pinyin, Nghĩa, Ví dụ VN, Ví dụ pinyin)
      hskData = rawData.slice(1).map(row => ({
        hanzi: row[0] || '',
        pinyin: row[1] || '',
        meaning: row[2] || '',
        exampleVn: row[3] || '',
        examplePinyin: row[4] || ''
      })).filter(item => item.hanzi);

      if (hskData.length > 0) {
        const savedIndex = localStorage.getItem('hsk1_current_index');
        if (savedIndex !== null) {
          currentIndex = parseInt(savedIndex, 10);
        }

        document.getElementById('total-cards').innerText = hskData.length;
        renderFlashcard();
        generateSentenceQuiz();
      }
    })
    .catch(err => console.error("Lỗi khi đọc file Excel:", err));
});

/* ==========================================
   1. LOGIC FLASHCARD
   ========================================== */
function renderFlashcard() {
  const item = hskData[currentIndex];

  document.getElementById('current-index').innerText = currentIndex + 1;
  document.getElementById('card-hanzi').innerText = item.hanzi;
  document.getElementById('card-pinyin').innerText = item.pinyin;
  document.getElementById('card-meaning').innerText = item.meaning;
  document.getElementById('card-example-cn').innerText = item.examplePinyin;
  document.getElementById('card-example-vn').innerText = item.exampleVn;

  const cardElement = document.getElementById('flashcard');
  if (cardElement) cardElement.classList.remove('flipped');

  localStorage.setItem('hsk1_current_index', currentIndex);
}

function flipCard() {
  document.getElementById('flashcard').classList.toggle('flipped');
}

function nextCard() {
  currentIndex = (currentIndex + 1) % hskData.length;
  renderFlashcard();
}

function prevCard() {
  currentIndex = (currentIndex - 1 + hskData.length) % hskData.length;
  renderFlashcard();
}

/* ==========================================
   2. LOGIC SẮP XẾP CÂU & PHÁO HOA
   ========================================== */
function generateSentenceQuiz() {
  document.getElementById('quiz-feedback').innerText = '';
  document.getElementById('user-sentence').innerHTML = '<span class="placeholder-text">Click chọn các từ bên dưới để ghép câu...</span>';
  userSelectedWords = [];

  const validItems = hskData.filter(item => item.examplePinyin && item.examplePinyin.trim().length > 0);
  if (validItems.length === 0) return;

  const selectedItem = validItems[Math.floor(Math.random() * validItems.length)];
  currentTargetSentence = selectedItem.examplePinyin.replace(/[。!？]/g, '').trim();
  
  document.getElementById('sentence-meaning').innerText = `"${selectedItem.exampleVn || selectedItem.meaning}"`;

  let rawWords = currentTargetSentence.split(' ').filter(w => w.trim() !== '');
  if (rawWords.length <= 1) {
    rawWords = currentTargetSentence.split('');
  }

  currentWords = [...rawWords].sort(() => Math.random() - 0.5);
  renderWordPool();
}

function renderWordPool() {
  const poolContainer = document.getElementById('word-pool');
  poolContainer.innerHTML = '';

  currentWords.forEach((word, index) => {
    const chip = document.createElement('button');
    chip.className = 'word-chip';
    chip.innerText = word;
    chip.onclick = () => selectWord(word, index, chip);
    poolContainer.appendChild(chip);
  });
}

function selectWord(word, index, btnElement) {
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
    checkSentenceAnswer();
  }
}

function resetSentence() {
  userSelectedWords = [];
  document.getElementById('user-sentence').innerHTML = '<span class="placeholder-text">Click chọn các từ bên dưới để ghép câu...</span>';
  document.getElementById('quiz-feedback').innerText = '';
  renderWordPool();
}

function checkSentenceAnswer() {
  const userResult = userSelectedWords.join(' ');
  const cleanTarget = currentTargetSentence.split(' ').join(' ');
  const feedback = document.getElementById('quiz-feedback');

  if (userResult === cleanTarget) {
    feedback.style.color = '#1f883d';
    feedback.innerText = '🎉 太棒了! Ghép câu chính xác!';

    if (typeof confetti === 'function') {
      confetti({
        particleCount: 100,
        spread: 70,
        origin: { y: 0.7 }
      });
    }

    setTimeout(generateSentenceQuiz, 2000);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Chưa đúng thứ tự! Bấm "Làm lại câu này" để thử lại.';
  }
}

/* ==========================================
   3. DARK MODE
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
