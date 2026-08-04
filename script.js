let hskData = [];

/* ==========================================
   BỘ TẠO ÂM THANH (WEB AUDIO API - NATIVE)
   ========================================== */
const AudioFX = {
  ctx: null,
  init() {
    if (!this.ctx) {
      this.ctx = new (window.AudioContext || window.webkitAudioContext)();
    }
  },
  // Âm thanh khi trả lời ĐÚNG (Ting vui tươi)
  playCorrect() {
    this.init();
    const now = this.ctx.currentTime;
    const osc = this.ctx.createOscillator();
    const gain = this.ctx.createGain();

    osc.type = 'sine';
    osc.frequency.setValueAtTime(523.25, now); // Nốt C5
    osc.frequency.exponentialRampToValueAtTime(880, now + 0.15); // Nốt A5

    gain.gain.setValueAtTime(0.3, now);
    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.3);

    osc.connect(gain);
    gain.connect(this.ctx.destination);

    osc.start(now);
    osc.stop(now + 0.3);
  },
  // Âm thanh khi trả lời SAI (Trầm nhẹ)
  playWrong() {
    this.init();
    const now = this.ctx.currentTime;
    const osc = this.ctx.createOscillator();
    const gain = this.ctx.createGain();

    osc.type = 'triangle';
    osc.frequency.setValueAtTime(180, now);
    osc.frequency.setValueAtTime(110, now + 0.1);

    gain.gain.setValueAtTime(0.3, now);
    gain.gain.exponentialRampToValueAtTime(0.01, now + 0.25);

    osc.connect(gain);
    gain.connect(this.ctx.destination);

    osc.start(now);
    osc.stop(now + 0.25);
  },
  // Âm thanh PHÁO HOA TỔNG KẾT (Chuỗi âm mừng chiến thắng)
  playCelebration() {
    this.init();
    const notes = [523.25, 659.25, 783.99, 1046.50]; // C5, E5, G5, C6
    notes.forEach((freq, index) => {
      setTimeout(() => {
        const now = this.ctx.currentTime;
        const osc = this.ctx.createOscillator();
        const gain = this.ctx.createGain();

        osc.type = 'sine';
        osc.frequency.setValueAtTime(freq, now);

        gain.gain.setValueAtTime(0.3, now);
        gain.gain.exponentialRampToValueAtTime(0.01, now + 0.3);

        osc.connect(gain);
        gain.connect(this.ctx.destination);

        osc.start(now);
        osc.stop(now + 0.3);
      }, index * 120);
    });
  }
};

/* ==========================================
   BỘ LỜI KHEN THÔNG MINH
   ========================================== */
const PraiseData = {
  high: [
    { title: "Bậc Thầy HSK 1! 🏆", msg: "Phong độ đỉnh cao! Cứ giữ đà này thì thi HSK 1 chỉ là chuyện nhỏ.", icon: "👑" },
    { title: "Xuất Sắc Vượt Trội! 🎉", msg: "Trí nhớ siêu đỉnh! Học viên xuất sắc nhất hôm nay chính là bạn.", icon: "🌟" }
  ],
  medium: [
    { title: "Nỗ Lực Rất Tốt! 💪", msg: "Tiến bộ rõ rệt qua từng phiên học. Cố gắng phát huy nhé!", icon: "🚀" },
    { title: "Làm Tốt Lắm! 👍", msg: "Chỉ cần ôn nhẹ lại vài câu chưa đúng là thuộc làu làu ngay.", icon: "⚡" }
  ],
  low: [
    { title: "Chiến Binh Kiên Trì! 🛡️", msg: "Vạn sự khởi đầu nan. Đừng lo lắng, sai đâu mình sửa đó nha!", icon: "🌱" },
    { title: "Không Sao Cả, Cố Lên! ❤️", msg: "Mỗi câu sai là một cơ hội để nhớ lâu hơn. Ôn lại kho câu sai nhé!", icon: "☀️" }
  ]
};

function getRandomPraise(accuracyRatio) {
  let list = PraiseData.low;
  if (accuracyRatio >= 0.8) list = PraiseData.high;
  else if (accuracyRatio >= 0.5) list = PraiseData.medium;

  return list[Math.floor(Math.random() * list.length)];
}

/* KỞI TẠO VÀ ĐỌC EXCEL */
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
        initBuilderApp();
      }
    })
    .catch(err => console.error("Lỗi đọc file Excel:", err));
});

function switchTab(tabId, btnElement) {
  document.querySelectorAll('.tab-content').forEach(tab => tab.classList.remove('active'));
  document.querySelectorAll('.tab-btn').forEach(btn => btn.classList.remove('active'));

  document.getElementById(tabId).classList.add('active');
  btnElement.classList.add('active');
}

/* ==========================================
   TAB 1: TRẮC NGHIỆM TỪ VỰNG
   ========================================== */
let mcList = [];
let mcGlobalIndex = 0;
let mcSessionList = [];
let mcSessionIndex = 0;
let mcCorrectFirstTry = 0;
let mcWrongList = JSON.parse(localStorage.getItem('hsk1_mc_wrong') || '[]');
let isMcWrongMode = false;
let currentMcItem = null;
let mcHadWrongInCurrent = false;

function initMultipleChoiceApp() {
  const savedOrder = localStorage.getItem('hsk1_mc_order');
  const savedIndex = localStorage.getItem('hsk1_mc_index');

  if (savedOrder) {
    const indices = JSON.parse(savedOrder);
    mcList = indices.map(idx => hskData[idx]).filter(Boolean);
  } else {
    const indices = hskData.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_mc_order', JSON.stringify(indices));
    mcList = indices.map(idx => hskData[idx]);
  }

  mcGlobalIndex = savedIndex ? parseInt(savedIndex, 10) : 0;
  if (mcGlobalIndex >= mcList.length) mcGlobalIndex = 0;

  updateMcSetupUI();
}

function updateMcSetupUI() {
  document.getElementById('mc-total-count').innerText = mcList.length;
  document.getElementById('mc-learned-count').innerText = mcGlobalIndex;
  document.getElementById('mc-remain-count').innerText = mcList.length - mcGlobalIndex;

  const wrongBtn = document.getElementById('mc-wrong-btn');
  if (mcWrongList.length > 0) {
    wrongBtn.style.display = 'inline-block';
    document.getElementById('mc-wrong-count').innerText = mcWrongList.length;
  } else {
    wrongBtn.style.display = 'none';
  }

  document.getElementById('mc-setup-screen').style.display = 'block';
  document.getElementById('mc-quiz-screen').style.display = 'none';
  document.getElementById('mc-summary-screen').style.display = 'none';
}

function startMcSession() {
  const countInput = parseInt(document.getElementById('mc-session-input').value, 10) || 10;
  
  if (mcGlobalIndex >= mcList.length) mcGlobalIndex = 0;

  mcSessionList = mcList.slice(mcGlobalIndex, mcGlobalIndex + countInput);
  if (mcSessionList.length === 0) {
    alert("Bạn đã học hết toàn bộ từ vựng! Hệ thống sẽ xoay vòng lại từ đầu.");
    mcGlobalIndex = 0;
    localStorage.setItem('hsk1_mc_index', 0);
    mcSessionList = mcList.slice(0, countInput);
  }

  mcSessionIndex = 0;
  mcCorrectFirstTry = 0;
  isMcWrongMode = false;

  document.getElementById('mc-setup-screen').style.display = 'none';
  document.getElementById('mc-quiz-screen').style.display = 'block';
  document.getElementById('mc-mode-badge').innerText = 'Phiên học mới';

  loadMcQuestion();
}

function startMcWrongReview() {
  if (mcWrongList.length === 0) return;
  
  mcSessionList = [...mcWrongList];
  mcSessionIndex = 0;
  mcCorrectFirstTry = 0;
  isMcWrongMode = true;

  document.getElementById('mc-setup-screen').style.display = 'none';
  document.getElementById('mc-quiz-screen').style.display = 'block';
  document.getElementById('mc-mode-badge').innerText = 'Ôn câu sai ⚠️';

  loadMcQuestion();
}

function loadMcQuestion() {
  document.getElementById('mc-feedback').innerText = '';

  if (mcSessionIndex >= mcSessionList.length) {
    showMcSummary();
    return;
  }

  currentMcItem = mcSessionList[mcSessionIndex];

  document.getElementById('mc-progress-text').innerText = `${mcSessionIndex + 1}/${mcSessionList.length}`;
  const pct = Math.round(((mcSessionIndex + 1) / mcSessionList.length) * 100);
  document.getElementById('mc-progress-bar').style.width = `${pct}%`;

  document.getElementById('mc-meaning').innerText = `"${currentMcItem.meaning}"`;

  const wrongOptions = hskData
    .filter(item => item.hanzi !== currentMcItem.hanzi)
    .sort(() => Math.random() - 0.5)
    .slice(0, 2);

  const options = [currentMcItem, ...wrongOptions].sort(() => Math.random() - 0.5);

  const container = document.getElementById('mc-options');
  container.innerHTML = '';

  options.forEach(opt => {
    const btn = document.createElement('button');
    btn.className = 'option-btn';
    btn.onclick = () => checkMultipleChoice(opt);
    btn.innerHTML = `
      <span class="option-hanzi">${opt.hanzi}</span>
      <span class="option-pinyin">${opt.pinyin}</span>
    `;
    container.appendChild(btn);
  });
}

function checkMultipleChoice(selected) {
  const feedback = document.getElementById('mc-feedback');

  if (selected.hanzi === currentMcItem.hanzi) {
    AudioFX.playCorrect(); // 🔊 Âm thanh đúng
    feedback.style.color = '#1f883d';
    feedback.innerText = '🎉 Chính xác!';
    if (typeof confetti === 'function') confetti({ particleCount: 60, spread: 50, origin: { y: 0.4 } });

    if (!mcHadWrongInCurrent) mcCorrectFirstTry++;

    if (isMcWrongMode) {
      mcWrongList = mcWrongList.filter(item => item.hanzi !== currentMcItem.hanzi);
      localStorage.setItem('hsk1_mc_wrong', JSON.stringify(mcWrongList));
    }

    setTimeout(() => {
      mcSessionIndex++;
      if (!isMcWrongMode) {
        mcGlobalIndex++;
        localStorage.setItem('hsk1_mc_index', mcGlobalIndex);
      }
      mcHadWrongInCurrent = false;
      loadMcQuestion();
    }, 1000);

  } else {
    AudioFX.playWrong(); // 🔊 Âm thanh sai
    feedback.style.color = '#d1242f';
    feedback.innerText = '❌ Chưa đúng, thử lại nhé!';
    mcHadWrongInCurrent = true;

    if (!mcWrongList.some(item => item.hanzi === currentMcItem.hanzi)) {
      mcWrongList.push(currentMcItem);
      localStorage.setItem('hsk1_mc_wrong', JSON.stringify(mcWrongList));
    }
  }
}

function skipMcQuestion() {
  mcSessionIndex++;
  if (!isMcWrongMode) {
    mcGlobalIndex++;
    localStorage.setItem('hsk1_mc_index', mcGlobalIndex);
  }
  mcHadWrongInCurrent = false;
  loadMcQuestion();
}

function showMcSummary() {
  document.getElementById('mc-quiz-screen').style.display = 'none';
  document.getElementById('mc-summary-screen').style.display = 'block';

  const total = mcSessionList.length;
  const correct = mcCorrectFirstTry;
  const accuracy = total > 0 ? correct / total : 0;

  document.getElementById('mc-sum-done').innerText = total;
  document.getElementById('mc-sum-correct').innerText = correct;
  document.getElementById('mc-sum-wrong').innerText = total - correct;

  // Lời khen
  const praise = getRandomPraise(accuracy);
  document.getElementById('mc-praise-title').innerText = praise.title;
  document.getElementById('mc-praise-message').innerText = praise.msg;
  document.getElementById('mc-badge-icon').innerText = praise.icon;

  AudioFX.playCelebration(); // 🔊 Âm pháo hoa
  if (typeof confetti === 'function') confetti({ particleCount: 100, spread: 70, origin: { y: 0.5 } });
}

function continueMcSession() { updateMcSetupUI(); }
function finishMcSession() { updateMcSetupUI(); }


/* ==========================================
   TAB 2: SẮP XẾP CÂU
   ========================================== */
let builderList = [];
let builderGlobalIndex = 0;
let builderSessionList = [];
let builderSessionIndex = 0;
let builderCorrectFirstTry = 0;
let builderWrongList = JSON.parse(localStorage.getItem('hsk1_builder_wrong') || '[]');
let isBuilderWrongMode = false;
let currentCorrectWords = [];
let userSelectedPool = [];
let builderHadWrongInCurrent = false;

function initBuilderApp() {
  const validItems = hskData.filter(item => item.examplePinyin && item.exampleVn);

  const savedOrder = localStorage.getItem('hsk1_builder_order');
  const savedIndex = localStorage.getItem('hsk1_builder_index');

  if (savedOrder) {
    const indices = JSON.parse(savedOrder);
    builderList = indices.map(idx => validItems[idx]).filter(Boolean);
  } else {
    const indices = validItems.map((_, idx) => idx).sort(() => Math.random() - 0.5);
    localStorage.setItem('hsk1_builder_order', JSON.stringify(indices));
    builderList = indices.map(idx => validItems[idx]);
  }

  builderGlobalIndex = savedIndex ? parseInt(savedIndex, 10) : 0;
  if (builderGlobalIndex >= builderList.length) builderGlobalIndex = 0;

  updateBuilderSetupUI();
}

function updateBuilderSetupUI() {
  document.getElementById('builder-total-count').innerText = builderList.length;
  document.getElementById('builder-learned-count').innerText = builderGlobalIndex;
  document.getElementById('builder-remain-count').innerText = builderList.length - builderGlobalIndex;

  const wrongBtn = document.getElementById('builder-wrong-btn');
  if (builderWrongList.length > 0) {
    wrongBtn.style.display = 'inline-block';
    document.getElementById('builder-wrong-count').innerText = builderWrongList.length;
  } else {
    wrongBtn.style.display = 'none';
  }

  document.getElementById('builder-setup-screen').style.display = 'block';
  document.getElementById('builder-quiz-screen').style.display = 'none';
  document.getElementById('builder-summary-screen').style.display = 'none';
}

function startBuilderSession() {
  const countInput = parseInt(document.getElementById('builder-session-input').value, 10) || 10;

  if (builderGlobalIndex >= builderList.length) builderGlobalIndex = 0;

  builderSessionList = builderList.slice(builderGlobalIndex, builderGlobalIndex + countInput);
  if (builderSessionList.length === 0) {
    alert("Bạn đã ghép hết kho câu ví dụ! Hệ thống sẽ xoay vòng lại.");
    builderGlobalIndex = 0;
    localStorage.setItem('hsk1_builder_index', 0);
    builderSessionList = builderList.slice(0, countInput);
  }

  builderSessionIndex = 0;
  builderCorrectFirstTry = 0;
  isBuilderWrongMode = false;

  document.getElementById('builder-setup-screen').style.display = 'none';
  document.getElementById('builder-quiz-screen').style.display = 'block';
  document.getElementById('builder-mode-badge').innerText = 'Phiên ghép câu';

  loadBuilderQuestion();
}

function startBuilderWrongReview() {
  if (builderWrongList.length === 0) return;

  builderSessionList = [...builderWrongList];
  builderSessionIndex = 0;
  builderCorrectFirstTry = 0;
  isBuilderWrongMode = true;

  document.getElementById('builder-setup-screen').style.display = 'none';
  document.getElementById('builder-quiz-screen').style.display = 'block';
  document.getElementById('builder-mode-badge').innerText = 'Ôn câu sai ⚠️';

  loadBuilderQuestion();
}

function loadBuilderQuestion() {
  document.getElementById('builder-feedback').innerText = '';

  if (builderSessionIndex >= builderSessionList.length) {
    showBuilderSummary();
    return;
  }

  const item = builderSessionList[builderSessionIndex];

  document.getElementById('builder-progress-text').innerText = `${builderSessionIndex + 1}/${builderSessionList.length}`;
  const pct = Math.round(((builderSessionIndex + 1) / builderSessionList.length) * 100);
  document.getElementById('builder-progress-bar').style.width = `${pct}%`;

  document.getElementById('builder-meaning').innerText = `"${item.exampleVn}"`;

  let rawPinyin = item.examplePinyin.replace(/[。!？,.]/g, '').trim();
  currentCorrectWords = rawPinyin.split(/\s+/);

  let wordObjects = currentCorrectWords.map((word, id) => ({ id, text: word }));
  let shuffledWords = [...wordObjects].sort(() => Math.random() - 0.5);

  userSelectedPool = [];
  renderBuilderUI(shuffledWords);
}

function renderBuilderUI(poolWords) {
  const box = document.getElementById('selected-words-box');
  const pool = document.getElementById('word-pool');

  box.innerHTML = '';
  if (userSelectedPool.length === 0) {
    box.innerHTML = `<p class="placeholder-text">Click vào từng từ bên dưới để ghép câu...</p>`;
  } else {
    userSelectedPool.forEach((item, idx) => {
      const chip = document.createElement('span');
      chip.className = 'word-chip selected';
      chip.innerText = item.text;
      chip.onclick = () => unselectWord(idx, poolWords);
      box.appendChild(chip);
    });
  }

  pool.innerHTML = '';
  poolWords.forEach((item) => {
    const chip = document.createElement('button');
    chip.className = 'word-chip';
    chip.innerText = item.text;
    chip.onclick = () => selectWord(item, poolWords);
    pool.appendChild(chip);
  });
}

function selectWord(item, currentPool) {
  userSelectedPool.push(item);
  const newPool = currentPool.filter(w => w.id !== item.id);
  renderBuilderUI(newPool);
}

function unselectWord(selectedIdx, currentPool) {
  const removed = userSelectedPool.splice(selectedIdx, 1)[0];
  currentPool.push(removed);
  renderBuilderUI(currentPool);
}

function resetBuilderSelection() {
  loadBuilderQuestion();
}

function checkBuilderAnswer() {
  const feedback = document.getElementById('builder-feedback');
  const userResult = userSelectedPool.map(w => w.text).join(' ');
  const correctResult = currentCorrectWords.join(' ');

  if (userSelectedPool.length === 0) {
    feedback.style.color = '#d1242f';
    feedback.innerText = '⚠️ Bạn chưa chọn từ nào kìa!';
    return;
  }

  const currentItem = builderSessionList[builderSessionIndex];

  if (userResult.toLowerCase() === correctResult.toLowerCase()) {
    AudioFX.playCorrect(); // 🔊 Âm thanh đúng
    feedback.style.color = '#1f883d';
    feedback.innerText = `🎉 Chuẩn đét! Đáp án: "${correctResult}"`;
    if (typeof confetti === 'function') confetti({ particleCount: 70, spread: 60, origin: { y: 0.8 } });

    if (!builderHadWrongInCurrent) builderCorrectFirstTry++;

    if (isBuilderWrongMode) {
      builderWrongList = builderWrongList.filter(item => item.exampleVn !== currentItem.exampleVn);
      localStorage.setItem('hsk1_builder_wrong', JSON.stringify(builderWrongList));
    }

    setTimeout(() => {
      builderSessionIndex++;
      if (!isBuilderWrongMode) {
        builderGlobalIndex++;
        localStorage.setItem('hsk1_builder_index', builderGlobalIndex);
      }
      builderHadWrongInCurrent = false;
      loadBuilderQuestion();
    }, 1200);

  } else {
    AudioFX.playWrong(); // 🔊 Âm thanh sai
    feedback.style.color = '#d1242f';
    feedback.innerText = `❌ Sai thứ tự từ rồi, thử lại nhé!`;
    builderHadWrongInCurrent = true;

    if (!builderWrongList.some(item => item.exampleVn === currentItem.exampleVn)) {
      builderWrongList.push(currentItem);
      localStorage.setItem('hsk1_builder_wrong', JSON.stringify(builderWrongList));
    }
  }
}

function skipBuilderSentence() {
  builderSessionIndex++;
  if (!isBuilderWrongMode) {
    builderGlobalIndex++;
    localStorage.setItem('hsk1_builder_index', builderGlobalIndex);
  }
  builderHadWrongInCurrent = false;
  loadBuilderQuestion();
}

function showBuilderSummary() {
  document.getElementById('builder-quiz-screen').style.display = 'none';
  document.getElementById('builder-summary-screen').style.display = 'block';

  const total = builderSessionList.length;
  const correct = builderCorrectFirstTry;
  const accuracy = total > 0 ? correct / total : 0;

  document.getElementById('builder-sum-done').innerText = total;
  document.getElementById('builder-sum-correct').innerText = correct;
  document.getElementById('builder-sum-wrong').innerText = total - correct;

  const praise = getRandomPraise(accuracy);
  document.getElementById('builder-praise-title').innerText = praise.title;
  document.getElementById('builder-praise-message').innerText = praise.msg;
  document.getElementById('builder-badge-icon').innerText = praise.icon;

  AudioFX.playCelebration(); // 🔊 Âm pháo hoa
  if (typeof confetti === 'function') confetti({ particleCount: 100, spread: 70, origin: { y: 0.5 } });
}

function continueBuilderSession() { updateBuilderSetupUI(); }
function finishBuilderSession() { updateBuilderSetupUI(); }

/* Dark Mode Toggle */
function toggleDarkMode() {
  document.body.classList.toggle('dark-mode');
}
