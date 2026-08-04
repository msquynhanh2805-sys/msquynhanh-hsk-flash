/**
 * HSK Flashcard Learner - Core Application Script
 * Architecture: ES6 Vanilla JS Object-Oriented Pattern
 */

class FlashcardApp {
    constructor() {
        // Master state
        this.cardsData = [];         // Dữ liệu gốc từ Excel
        this.filteredCards = [];     // Dữ liệu đã qua lọc/tìm kiếm
        this.currentIndex = 0;
        this.isFlipped = false;

        // LocalStorage States
        this.userState = {
            rememberedIds: new Set(),
            forgotIds: new Set(),
            favoriteIds: new Set(),
            theme: 'light'
        };

        // Cache DOM elements
        this.dom = {
            themeToggle: document.getElementById('theme-toggle'),
            statTotal: document.getElementById('stat-total'),
            statRemembered: document.getElementById('stat-remembered'),
            statForgot: document.getElementById('stat-forgot'),
            statPercentage: document.getElementById('stat-percentage'),
            searchInput: document.getElementById('search-input'),
            filterSelect: document.getElementById('filter-select'),
            btnShuffle: document.getElementById('btn-shuffle'),
            progressBar: document.getElementById('learning-progress'),
            currentIndexDisplay: document.getElementById('current-index'),
            totalFilteredDisplay: document.getElementById('total-filtered-cards'),
            flashcard: document.getElementById('flashcard'),
            favBtn: document.getElementById('fav-btn'),
            cardHanzi: document.getElementById('card-hanzi'),
            cardPinyin: document.getElementById('card-pinyin'),
            cardMeaning: document.getElementById('card-meaning'),
            cardExamplePinyin: document.getElementById('card-example-pinyin'),
            cardExampleVn: document.getElementById('card-example-vn'),
            btnPrev: document.getElementById('btn-prev'),
            btnNext: document.getElementById('btn-next'),
            btnFlip: document.getElementById('btn-flip'),
            btnForgot: document.getElementById('btn-forgot'),
            btnRemembered: document.getElementById('btn-remembered'),
            modalCongrats: document.getElementById('modal-congrats'),
            btnRestart: document.getElementById('btn-restart')
        };

        this.init();
    }

    /**
     * Khởi tạo ứng dụng
     */
    async init() {
        this.loadStorage();
        this.applyTheme();
        this.bindEvents();
        await this.loadExcelData();
    }

    /**
     * Nạp dữ liệu LocalStorage
     */
    loadStorage() {
        const saved = localStorage.getItem('HSK_FLASHCARD_PROGRESS');
        if (saved) {
            try {
                const parsed = JSON.parse(saved);
                this.userState.rememberedIds = new Set(parsed.rememberedIds || []);
                this.userState.forgotIds = new Set(parsed.forgotIds || []);
                this.userState.favoriteIds = new Set(parsed.favoriteIds || []);
                this.userState.theme = parsed.theme || 'light';
            } catch (e) {
                console.error("Lỗi parse LocalStorage:", e);
            }
        }
    }

    /**
     * Lưu trạng thái vào LocalStorage
     */
    saveStorage() {
        const payload = {
            rememberedIds: Array.from(this.userState.rememberedIds),
            forgotIds: Array.from(this.userState.forgotIds),
            favoriteIds: Array.from(this.userState.favoriteIds),
            theme: this.userState.theme
        };
        localStorage.setItem('HSK_FLASHCARD_PROGRESS', JSON.stringify(payload));
    }

    /**
     * Nạp và Parse file Excel bằng SheetJS
     */
    async loadExcelData() {
        try {
            const response = await fetch('HSK1_flashcards.xlsx');
            if (!response.ok) throw new Error('Không tìm thấy file HSK1_flashcards.xlsx');
            
            const arrayBuffer = await response.arrayBuffer();
            const workbook = XLSX.read(arrayBuffer, { type: 'array' });
            
            // Lấy sheet đầu tiên
            const firstSheetName = workbook.SheetNames[0];
            const worksheet = workbook.Sheets[firstSheetName];
            
            // Chuyển dữ liệu sang dạng JSON
            const rawData = XLSX.utils.sheet_to_json(worksheet);

            // Mapping chuẩn hóa dữ liệu & gán Unique ID
            this.cardsData = rawData.map((item, index) => ({
                id: `word_${index}`,
                hanzi: item['Hanzi'] || item['hanzi'] || '',
                pinyin: item['Pinyin'] || item['pinyin'] || '',
                meaning: item['Nghĩa'] || item['nghia'] || item['Meaning'] || '',
                exampleVn: item['Ví dụ VN'] || item['vi_du_vn'] || '',
                examplePinyin: item['Ví dụ pinyin'] || item['vi_du_pinyin'] || ''
            }));

            this.applyFilter();
            this.updateDashboardStats();

        } catch (error) {
            console.error('Lỗi khi tải dữ liệu Excel:', error);
            this.dom.cardHanzi.textContent = "Lỗi tải File Excel";
            this.dom.cardMeaning.textContent = "Đảm bảo file HSK1_flashcards.xlsx nằm cùng thư mục.";
        }
    }

    /**
     * Lọc và Tìm kiếm thẻ
     */
    applyFilter() {
        const filterValue = this.dom.filterSelect.value;
        const searchQuery = this.dom.searchInput.value.toLowerCase().trim();

        this.filteredCards = this.cardsData.filter(card => {
            // Check bộ lọc
            let matchFilter = true;
            if (filterValue === 'remembered') matchFilter = this.userState.rememberedIds.has(card.id);
            else if (filterValue === 'forgot') matchFilter = this.userState.forgotIds.has(card.id);
            else if (filterValue === 'favorite') matchFilter = this.userState.favoriteIds.has(card.id);

            // Check tìm kiếm
            let matchSearch = true;
            if (searchQuery) {
                matchSearch = card.hanzi.toLowerCase().includes(searchQuery) ||
                              card.pinyin.toLowerCase().includes(searchQuery) ||
                              card.meaning.toLowerCase().includes(searchQuery);
            }

            return matchFilter && matchSearch;
        });

        this.currentIndex = 0;
        this.renderCurrentCard();
    }

    /**
     * Hiển thị Card hiện tại lên UI
     */
    renderCurrentCard() {
        this.resetCardFlip();

        if (this.filteredCards.length === 0) {
            this.dom.cardHanzi.textContent = "Không có từ";
            this.dom.cardPinyin.textContent = "";
            this.dom.cardMeaning.textContent = "Không tìm thấy dữ liệu phù hợp";
            this.dom.cardExamplePinyin.textContent = "";
            this.dom.cardExampleVn.textContent = "";
            this.dom.currentIndexDisplay.textContent = "0";
            this.dom.totalFilteredDisplay.textContent = "0";
            this.dom.favBtn.classList.remove('active');
            this.updateProgressBar();
            return;
        }

        const card = this.filteredCards[this.currentIndex];

        // Render mặt trước & sau
        this.dom.cardHanzi.textContent = card.hanzi;
        this.dom.cardPinyin.textContent = card.pinyin;
        this.dom.cardMeaning.textContent = card.meaning;
        this.dom.cardExamplePinyin.textContent = card.examplePinyin;
        this.dom.cardExampleVn.textContent = card.exampleVn;

        // Render chỉ số
        this.dom.currentIndexDisplay.textContent = this.currentIndex + 1;
        this.dom.totalFilteredDisplay.textContent = this.filteredCards.length;

        // Trạng thái Favorite
        if (this.userState.favoriteIds.has(card.id)) {
            this.dom.favBtn.classList.add('active');
            this.dom.favBtn.querySelector('i').className = 'fa-solid fa-star';
        } else {
            this.dom.favBtn.classList.remove('active');
            this.dom.favBtn.querySelector('i').className = 'fa-regular fa-star';
        }

        this.updateProgressBar();
    }

    /**
     * Cập nhật Dashboard Thống kê
     */
    updateDashboardStats() {
        const total = this.cardsData.length;
        const remembered = this.userState.rememberedIds.size;
        const forgot = this.userState.forgotIds.size;
        const percentage = total > 0 ? Math.round((remembered / total) * 100) : 0;

        this.dom.statTotal.textContent = total;
        this.dom.statRemembered.textContent = remembered;
        this.dom.statForgot.textContent = forgot;
        this.dom.statPercentage.textContent = `${percentage}%`;
    }

    /**
     * Cập nhật Thanh Tiến độ học tập
     */
    updateProgressBar() {
        if (this.filteredCards.length === 0) {
            this.dom.progressBar.style.width = `0%`;
            return;
        }
        const progress = ((this.currentIndex + 1) / this.filteredCards.length) * 100;
        this.dom.progressBar.style.width = `${progress}%`;
    }

    /**
     * Lật mặt thẻ 3D
     */
    toggleFlip() {
        if (this.filteredCards.length === 0) return;
        this.isFlipped = !this.isFlipped;
        this.dom.flashcard.classList.toggle('flipped', this.isFlipped);
    }

    /**
     * Đưa mặt thẻ về ban đầu
     */
    resetCardFlip() {
        this.isFlipped = false;
        this.dom.flashcard.classList.remove('flipped');
    }

    /**
     * Chuyển Card tiếp theo
     */
    nextCard() {
        if (this.currentIndex < this.filteredCards.length - 1) {
            this.currentIndex++;
            this.renderCurrentCard();
        } else if (this.filteredCards.length > 0) {
            this.showCongratsModal();
        }
    }

    /**
     * Chuyển Card trước đó
     */
    prevCard() {
        if (this.currentIndex > 0) {
            this.currentIndex--;
            this.renderCurrentCard();
        }
    }

    /**
     * Xử lý Đã nhớ
     */
    markRemembered() {
        if (this.filteredCards.length === 0) return;
        const currentCard = this.filteredCards[this.currentIndex];

        this.userState.rememberedIds.add(currentCard.id);
        this.userState.forgotIds.delete(currentCard.id);

        this.saveStorage();
        this.updateDashboardStats();
        this.nextCard();
    }

    /**
     * Xử lý Chưa nhớ -> Đẩy xuống cuối danh sách
     */
    markForgot() {
        if (this.filteredCards.length === 0) return;
        const currentCard = this.filteredCards[this.currentIndex];

        this.userState.forgotIds.add(currentCard.id);
        this.userState.rememberedIds.delete(currentCard.id);

        // Đẩy từ này xuống cuối mảng đã lọc
        this.filteredCards.splice(this.currentIndex, 1);
        this.filteredCards.push(currentCard);

        this.saveStorage();
        this.updateDashboardStats();

        // Giữ nguyên chỉ số index nhưng render dữ liệu mới vừa đè vào
        if (this.currentIndex >= this.filteredCards.length) {
            this.currentIndex = 0;
        }
        this.renderCurrentCard();
    }

    /**
     * Bật/Tắt từ khó
     */
    toggleFavorite(e) {
        e.stopPropagation(); // Không kích hoạt event lật thẻ
        if (this.filteredCards.length === 0) return;

        const currentCard = this.filteredCards[this.currentIndex];
        if (this.userState.favoriteIds.has(currentCard.id)) {
            this.userState.favoriteIds.delete(currentCard.id);
        } else {
            this.userState.favoriteIds.add(currentCard.id);
        }

        this.saveStorage();
        this.renderCurrentCard();
    }

    /**
     * Học ngẫu nhiên (Fisher-Yates Shuffle)
     */
    shuffleCards() {
        for (let i = this.filteredCards.length - 1; i > 0; i--) {
            const j = Math.floor(Math.random() * (i + 1));
            [this.filteredCards[i], this.filteredCards[j]] = [this.filteredCards[j], this.filteredCards[i]];
        }
        this.currentIndex = 0;
        this.renderCurrentCard();
    }

    /**
     * Đổi Dark/Light Mode
     */
    toggleTheme() {
        this.userState.theme = this.userState.theme === 'light' ? 'dark' : 'light';
        this.applyTheme();
        this.saveStorage();
    }

    applyTheme() {
        document.documentElement.setAttribute('data-theme', this.userState.theme);
        const icon = this.dom.themeToggle.querySelector('i');
        if (this.userState.theme === 'dark') {
            icon.className = 'fa-solid fa-sun';
        } else {
            icon.className = 'fa-solid fa-moon';
        }
    }

    /**
     * Hiện Modal chúc mừng
     */
    showCongratsModal() {
        this.dom.modalCongrats.classList.add('active');
    }

    hideCongratsModal() {
        this.dom.modalCongrats.classList.remove('active');
        this.currentIndex = 0;
        this.renderCurrentCard();
    }

    /**
     * Đăng ký Event Listeners
     */
    bindEvents() {
        // Lật thẻ
        this.dom.flashcard.addEventListener('click', () => this.toggleFlip());
        this.dom.btnFlip.addEventListener('click', () => this.toggleFlip());

        // Điều hướng & Đánh giá
        this.dom.btnNext.addEventListener('click', () => this.nextCard());
        this.dom.btnPrev.addEventListener('click', () => this.prevCard());
        this.dom.btnRemembered.addEventListener('click', () => this.markRemembered());
        this.dom.btnForgot.addEventListener('click', () => this.markForgot());
        this.dom.favBtn.addEventListener('click', (e) => this.toggleFavorite(e));

        // Controls
        this.dom.btnShuffle.addEventListener('click', () => this.shuffleCards());
        this.dom.filterSelect.addEventListener('change', () => this.applyFilter());
        this.dom.searchInput.addEventListener('input', () => this.applyFilter());
        this.dom.themeToggle.addEventListener('click', () => this.toggleTheme());
        this.dom.btnRestart.addEventListener('click', () => this.hideCongratsModal());

        // Phím tắt bàn phím (Keyboard Shortcuts)
        document.addEventListener('keydown', (e) => {
            if (e.target.tagName === 'INPUT') return; // Không ăn phím tắt khi search
            
            switch (e.code) {
                case 'Space':
                    e.preventDefault();
                    this.toggleFlip();
                    break;
                case 'ArrowRight':
                    this.nextCard();
                    break;
                case 'ArrowLeft':
                    this.prevCard();
                    break;
                case 'Key1':
                case 'Numpad1':
                    this.markForgot();
                    break;
                case 'Key2':
                case 'Numpad2':
                    this.markRemembered();
                    break;
            }
        });
    }
}

// Khởi chạy ứng dụng khi DOM sẵn sàng
document.addEventListener('DOMContentLoaded', () => {
    window.app = new FlashcardApp();
});
