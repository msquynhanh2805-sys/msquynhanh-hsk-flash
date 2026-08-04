let cards = [];
let current = 0;
let showingFront = true;

const STORAGE_KEY = "flashcard-progress";

fetch("data/HSK1_flashcards.xlsx")
    .then(res => res.arrayBuffer())
    .then(buffer => {

        const workbook = XLSX.read(buffer);

        const sheet = workbook.Sheets[workbook.SheetNames[0]];

        cards = XLSX.utils.sheet_to_json(sheet);

        loadProgress();

        showCard();

    });

function showCard() {

    if (current >= cards.length) {

        alert("🎉 Bạn đã học hết!");

        localStorage.removeItem(STORAGE_KEY);

        return;
    }

    let card = cards[current];

    document.getElementById("front").innerHTML = card["Chinese"];

    document.getElementById("back").innerHTML =
        "<b>" + card["Pinyin"] + "</b><br><br>" +
        card["Meaning"];

    document.getElementById("back").style.display = "none";

    showingFront = true;

    updateProgress();

}

document.getElementById("flipBtn").onclick = function () {

    if (showingFront) {

        document.getElementById("back").style.display = "block";

    } else {

        document.getElementById("back").style.display = "none";

    }

    showingFront = !showingFront;

}

document.getElementById("knowBtn").onclick = function () {

    current++;

    saveProgress();

    showCard();

}

document.getElementById("dontKnowBtn").onclick = function () {

    cards.push(cards[current]);

    current++;

    saveProgress();

    showCard();

}

function updateProgress() {

    let percent = current / cards.length * 100;

    document.getElementById("progressBar").style.width = percent + "%";

    document.getElementById("progressText").innerHTML =
        current + " / " + cards.length;

}

function saveProgress() {

    localStorage.setItem(STORAGE_KEY, current);

}

function loadProgress() {

    let saved = localStorage.getItem(STORAGE_KEY);

    if (saved) {

        current = parseInt(saved);

    }

}
