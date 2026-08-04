// Dữ liệu chung từ file Excel
let hskData = [];
let currentIndex = 0;
let currentQuizQuestion = null;

// Khởi tạo app khi trang web tải xong
window.addEventListener('DOMContentLoaded', () => {
  fetch('HSK1_flashcards.xlsx')
    .then(res => res.arrayBuffer())
    .then(buffer => {
      const workbook = XLSX.read(buffer, { type: 'array' });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const rawData = XLSX.utils.sheet_to_json(sheet, { header: 1 });

      // Lấy dữ liệu từ dòng thứ 2 (bỏ qua header)
      hskData = rawData.slice(1).map(row => ({
        hanzi: row[0] || '',
        pinyin: row[1] || '',
        meaning: row[2] || '',
        exampleVn: row[3] || '',
        examplePinyin: row[4] || ''
      })).filter(item => item.hanzi);

      if (hskData.length > 0) {
        // Tải lại tiến độ học cũ từ LocalStorage (nếu có)
        const savedIndex = localStorage.getItem('hsk1_current_index');
        if (savedIndex !== null) {
          currentIndex = parseInt(savedIndex, 10);
        }

        // Cập nhật giao diện
        document.getElementById('total-cards').innerText = hskData.length;
        renderFlashcard();
        generateQuiz();
      }
    })
    .catch(err => console.error("Lỗi đọc file Excel:", err));
});

/* ==========================================
   1. LOGIC FLASHCARD (MẶT TRƯỚC / MẶT SAU)
   ========================================== */
function renderFlashcard() {
  const item = hskData[currentIndex];
  
  // Hiển thị số thứ tự
  document.getElementById('current-index').innerText = currentIndex + 1;
  
  // Điền dữ liệu vào mặt trước/sau của thẻ
  document.getElementById('card-hanzi').innerText = item.hanzi;
  document.getElementById('card-pinyin').innerText = item.pinyin;
  document.getElementById('card-meaning').innerText = item.meaning;
  document.getElementById('card-example-cn').innerText = item.examplePinyin;
  document.getElementById('card-example-vn').innerText = item.exampleVn;

  // Trở về mặt trước khi chuyển bài
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
   2. LOGIC BÀI TẬP ÔN TẬP (QUIZ PHÁO HOA)
   ========================================== */
function generateQuiz() {
  const feedback = document.getElementById('quiz-feedback');
  if (feedback) feedback.innerText = '';

  // Bốc ngẫu nhiên 1 từ làm câu hỏi
  const randIdx = Math.floor(Math.random() * hskData.length);
  currentQuizQuestion = hskData[randIdx];

  document.getElementById('quiz-meaning').innerText = `"${currentQuizQuestion.meaning}"`;

  // Tạo 2 đáp án sai để làm nhiễu
  let options = [currentQuizQuestion];
  while (options.length < 3 && options.length < hskData.length) {
    let wrong = hskData[Math.floor(Math.random() * hskData.length)];
    if (!options.some(opt => opt.hanzi === wrong.hanzi)) {
      options.push(wrong);
    }
  }

  // Trộn thứ tự đáp án
  options.sort(() => Math.random() - 0.5);

  // Hiển thị nút bấm
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
    feedback.style.color = '#2ea043';
    feedback.innerText = '🎉 太棒了! Bán pháo hoa ăn mừng!';

    // 🎆 BẮN PHÁO HOA
    if (typeof confetti === 'function') {
      confetti({ particleCount: 90, spread: 70, origin: { y: 0.7 } });
    }

    // Đổi câu hỏi sau 1.5s
    setTimeout(generateQuiz, 1500);
  } else {
    feedback.style.color = '#da3633';
    feedback.innerText = `❌ Sai rồi! "${currentQuizQuestion.meaning}" là: ${currentQuizQuestion.hanzi} (${currentQuizQuestion.pinyin})`;
  }
}

/* ==========================================
   3. CHẾ ĐỘ SÁNG / TỐI (DARK MODE)
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
