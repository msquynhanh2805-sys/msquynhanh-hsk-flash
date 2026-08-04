let hskData = [];

// Khởi tạo tính năng nhận diện giọng nói (Speech-to-Text)
const SpeechRecognition = window.SpeechRecognition || window.webkitSpeechRecognition;
let recognition = null;

if (SpeechRecognition) {
  recognition = new SpeechRecognition();
  recognition.lang = 'zh-CN'; // Cấu hình nhận diện tiếng Trung Quốc (Phổ thông)
  recognition.interimResults = false;
  recognition.maxAlternatives = 1;
}

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
        initMultipleChoiceApp();
        initSpeakingApp();
      }
    })
    .catch(err => console.error("Lỗi đọc file Excel:", err));
});

/* ==========================================
   DẠNG 1: TRẮC NGHIỆM TỪ VỰNG (LƯU TIẾN ĐỘ)
   ========================================== */
let mcList = [];
let mcIndex = 0;
let currentMcCorrectItem = null;

function initMultipleChoiceApp() {
  const savedMcOrder = localStorage.getItem('hsk1_mc_order');
  const savedMcIndex = localStorage.getItem('hsk1_mc_index');

  if (savedMcOrder) {
    const indices = JSON.parse(savedMcOrder);
    mcList = indices.map(idx => hskData[idx]).filter(Boolean);
  } else {
    const indices = hskData.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_mc_order', JSON.stringify(indices));
    mcList = indices.map(idx => hskData[idx]);
  }

  if (savedMcIndex) {
    mcIndex = parseInt(savedMcIndex, 10);
    if (mcIndex >= mcList.length) mcIndex = 0;
  }

  loadMcQuestion();
}

function loadMcQuestion() {
  document.getElementById('mc-feedback').innerText = '';
  document.getElementById('mc-progress').innerText = `Tiến độ từ: ${mcIndex + 1} / ${mcList.length}`;

  currentMcCorrectItem = mcList[mcIndex];
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

    mcIndex++;
    localStorage.setItem('hsk1_mc_index', mcIndex);
    setTimeout(loadMcQuestion, 1200);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Sai rồi, thử lại xem!';
  }
}

function skipMcQuestion() {
  mcIndex++;
  localStorage.setItem('hsk1_mc_index', mcIndex);
  loadMcQuestion();
}

/* ==========================================
   DẠNG 2: LUYỆN NÓI CÂU (CHE PINYIN HÒAN TOÀN)
   ========================================== */
let speakingList = [];
let currentSpeakingIndex = 0;
let currentTargetPinyin = "";

function initSpeakingApp() {
  const validItems = hskData.filter(item => item.examplePinyin && item.exampleVn);

  const savedOrder = localStorage.getItem('hsk1_speaking_order');
  const savedIndex = localStorage.getItem('hsk1_speaking_index');

  if (savedOrder) {
    const indices = JSON.parse(savedOrder);
    speakingList = indices.map(idx => validItems[idx]).filter(Boolean);
  } else {
    const indices = validItems.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_speaking_order', JSON.stringify(indices));
    speakingList = indices.map(idx => validItems[idx]);
  }

  if (savedIndex) {
    currentSpeakingIndex = parseInt(savedIndex, 10);
    if (currentSpeakingIndex >= speakingList.length) currentSpeakingIndex = 0;
  }

  loadSpeakingQuestion();
}

function loadSpeakingQuestion() {
  document.getElementById('speaking-feedback').innerText = '';
  document.getElementById('speaking-progress').innerText = `Tiến độ câu: ${currentSpeakingIndex + 1} / ${speakingList.length}`;

  const item = speakingList[currentSpeakingIndex];
  document.getElementById('speaking-meaning').innerText = `"${item.exampleVn}"`;
  
  currentTargetPinyin = item.examplePinyin.replace(/[。!？,.]/g, '').trim();
  
  // CHE PINYIN HOÀN TOÀN
  const pinyinHintElement = document.getElementById('speaking-pinyin-hint');
  pinyinHintElement.innerText = "🙈 * * * * * * *";
  pinyinHintElement.style.color = "#8c959f";
}

function showPinyinHint() {
  const pinyinHintElement = document.getElementById('speaking-pinyin-hint');
  pinyinHintElement.innerText = currentTargetPinyin;
  pinyinHintElement.style.color = "#0969da";
}

function startListening() {
  const feedback = document.getElementById('speaking-feedback');

  if (!recognition) {
    feedback.style.color = '#d1242f';
    feedback.innerText = '⚠️ Trình duyệt không hỗ trợ micro (Hãy mở bằng Chrome hoặc Safari nhé!).';
    return;
  }

  feedback.style.color = '#0969da';
  feedback.innerText = '🎙️ Đang nghe... M hãy đọc câu tiếng Trung đi!';

  try {
    recognition.start();
  } catch (e) {
    recognition.stop();
    recognition.start();
  }

  recognition.onresult = (event) => {
    const spokenText = event.results[0][0].transcript.replace(/[。!？,.]/g, '').trim();
    checkSpeakingAnswer(spokenText);
  };

  recognition.onerror = (event) => {
    feedback.style.color = '#d1242f';
    if (event.error === 'not-allowed') {
      feedback.innerText = '⚠️ Trình duyệt chưa được cấp quyền dùng Micro!';
    } else {
      feedback.innerText = '❌ Bấm mic thử đọc lại lần nữa nhé!';
    }
  };
}

function checkSpeakingAnswer(spokenText) {
  const feedback = document.getElementById('speaking-feedback');

  // HIỂN THỊ PINYIN ĐÁP ÁN KHI NÓI XONG
  showPinyinHint();

  if (spokenText && spokenText.length > 0) {
    feedback.style.color = '#1f883d';
    feedback.innerText = `🎉 Đỉnh quá! App nghe được: "${spokenText}"`;
    if (typeof confetti === 'function') confetti({ particleCount: 90, spread: 70, origin: { y: 0.8 } });

    currentSpeakingIndex++;
    localStorage.setItem('hsk1_speaking_index', currentSpeakingIndex);
    setTimeout(loadSpeakingQuestion, 2200);
  } else {
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Chưa nghe rõ! Thử phát âm to và rõ ràng hơn nhé.`;
  }
}

// PHÁT ÂM CHUẨN GIỌNG NÓI TIẾNG TRUNG
function speakSample() {
  const item = speakingList[currentSpeakingIndex];
  if (!item || !item.examplePinyin) return;

  // Lấy thẳng Hán tự hoặc Pinyin phát âm qua SpeechSynthesis giọng chuẩn zh-CN
  if ('speechSynthesis' in window) {
    window.speechSynthesis.cancel();
    
    // Ưu tiên đọc câu Hán tự nếu có, không có thì đọc Pinyin
    const text = item.hanziExample || item.examplePinyin; 
    const utterance = new SpeechSynthesisUtterance(text);
    utterance.lang = 'zh-CN';
    utterance.rate = 0.85; // Tốc độ chuẩn cho người học

    // Chọn giọng tiếng Trung chuẩn của trình duyệt
    const voices = window.speechSynthesis.getVoices();
    const zhVoice = voices.find(v => v.lang.includes('zh') || v.lang.includes('CN'));
    if (zhVoice) utterance.voice = zhVoice;

    window.speechSynthesis.speak(utterance);
  }
}

function skipSpeakingSentence() {
  currentSpeakingIndex++;
  localStorage.setItem('hsk1_speaking_index', currentSpeakingIndex);
  loadSpeakingQuestion();
}

/* ==========================================
   DARK MODE
   ========================================== */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
