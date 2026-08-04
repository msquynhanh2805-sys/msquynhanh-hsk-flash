let cards = [];
let current = 0;

const STORAGE_KEY = "hsk_flashcards_progress";

const card = document.getElementById("card");
const progressBar = document.getElementById("progressBar");
const progressText = document.getElementById("progressText");

const hanzi = document.getElementById("hanzi");
const pinyin = document.getElementById("pinyin");
const meaning = document.getElementById("meaning");
const exampleVN = document.getElementById("exampleVN");
const examplePY = document.getElementById("examplePY");

// Đọc file Excel ở cùng thư mục với index.html
fetch("HSK1_flashcards.xlsx")
.then(res => {
    if (!res.ok) throw new Error("Không tìm thấy file Excel");
    return res.arrayBuffer();
})
.then(buffer => {

    const workbook = XLSX.read(buffer);

    const sheet = workbook.Sheets[workbook.SheetNames[0]];

    cards = XLSX.utils.sheet_to_json(sheet);

    loadProgress();

    showCard();

})
.catch(err=>{
    alert("Không đọc được file Excel. Hãy kiểm tra tên file là HSK1_flashcards.xlsx");
    console.error(err);
});

function showCard(){

    if(cards.length===0) return;

    if(current>=cards.length){

        alert("🎉 Chúc mừng! Bạn đã học xong.");

        localStorage.removeItem(STORAGE_KEY);

        current=0;

    }

    let c=cards[current];

    hanzi.textContent=c["Hanzi"]||"";

    pinyin.textContent=c["Pinyin"]||"";

    meaning.textContent=c["Nghĩa"]||"";

    exampleVN.textContent=c["Ví dụ VN"]||"";

    examplePY.textContent=c["Ví dụ pinyin"]||"";

    card.classList.remove("flipped");

    updateProgress();

}

function updateProgress(){

    progressBar.style.width=((current+1)/cards.length*100)+"%";

    progressText.innerHTML=(current+1)+" / "+cards.length;

}

document.getElementById("flipBtn").onclick=()=>{

    card.classList.toggle("flipped");

}

card.onclick=()=>{

    card.classList.toggle("flipped");

}

document.getElementById("knowBtn").onclick=()=>{

    current++;

    saveProgress();

    showCard();

}

document.getElementById("dontKnowBtn").onclick=()=>{

    cards.push(cards[current]);

    current++;

    saveProgress();

    showCard();

}

function saveProgress(){

    localStorage.setItem(STORAGE_KEY,current);

}

function loadProgress(){

    const saved=localStorage.getItem(STORAGE_KEY);

    if(saved){

        current=parseInt(saved);

    }

}
