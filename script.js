let hskData = [];
let currentIndex = 0;
let currentQuizQuestion = null;

// Khởi tạo app khi đọc xong file Excel HSK1_flashcards.xlsx
window.addEventListener('DOMContentLoaded', () => {
  fetch('HSK1_flashcards.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Lấy từ dòng thứ 2 (bỏ qua header 5 cột)
      hskData = rawData.slice(1).map(row => ({
        hanzi: row[0] || '',
        pinyin: row[1] || '',
        meaning: row[2] || '',
        exampleVn: row[3] || '',
        examplePinyin: row[4] || ''
      })).filter(item => item.hanzi);

      if (hskData.length > 0) {
        // Tải tiến độ cũ từ LocalStorage nếu có
        const savedIndex = localStorage.getItem('hsk1_current_index');
        if (savedIndex !== null) {
          currentIndex = parseInt(savedIndex, 10);
        }

        document.getElementById('total-cards').innerText = hskData.length;
        renderFlashcard();
        generateQuiz();
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

  // Trở về mặt trước
  const cardElement = document.getElementById('flashcard');
  if (cardElement) cardElement.classList.remove('flipped');

  // Lưu tiến độ vào LocalStorage
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
   2. LOGIC QUIZ PHÁO HOA
   ========================================== */
function generateQuiz() {
  const feedback = document.getElementById('quiz-feedback');
  if (feedback) feedback.innerText = '';

  // 1. Chọn 1 từ ngẫu nhiên làm câu hỏi
  const randIdx = Math.floor(Math.random() * hskData.length);
  currentQuizQuestion = hskData[randIdx];

  document.getElementById('quiz-meaning').innerText = `"${currentQuizQuestion.meaning}"`;

  // 2. Tạo 2 đáp án sai ngẫu nhiên
  let options = [currentQuizQuestion];
  while (options.length < 3 && options.length < hskData.length) {
    let wrong = hskData[Math.floor(Math.random() * hskData.length)];
    if (!options.some(opt => opt.hanzi === wrong.hanzi)) {
      options.push(wrong);
    }
  }

  // 3. Trộn ngẫu nhiên thứ tự đáp án
  options.sort(() => Math.random() - 0.5);

  // 4. Render các nút bấm đáp án
  const container = document.getElementById('quiz-options');
  container.innerHTML = '';

  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'quiz-btn';
    btn.innerHTML = `
      <span class="opt-hanzi">${opt.hanzi}</span>
      <span class="opt-pinyin">${opt.pinyin}</span>
    `;
    btn.onclick = () => checkQuiz(opt);
    container.appendChild(btn);
  });
}

function checkQuiz(selected) {
  const feedback = document.getElementById('quiz-feedback');

  if (selected.hanzi === currentQuizQuestion.hanzi) {
    feedback.style.color = '#1f883d';
    feedback.innerText = '🎉 太棒了! Chính xác!';

    // BẮN PHÁO HOA 🎆
    if (typeof confetti === 'function') {
      confetti({
        particleCount: 80,
        spread: 70,
        origin: { y: 0.7 }
      });
    }

    // Tự động nhảy sang câu hỏi tiếp theo sau 1.5 giây
    setTimeout(generateQuiz, 1500);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Chưa đúng! "${currentQuizQuestion.meaning}" là: ${currentQuizQuestion.hanzi} (${currentQuizQuestion.pinyin})`;
  }
}

/* ==========================================
   3. DARK MODE
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
