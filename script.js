*{
    margin:0;
    padding:0;
    box-sizing:border-box;
}

body{
    font-family:Arial,Helvetica,sans-serif;
    background:linear-gradient(135deg,#dbeafe,#eff6ff);
    min-height:100vh;
    display:flex;
    justify-content:center;
    align-items:center;
    padding:20px;
}

.container{
    width:100%;
    max-width:500px;
    text-align:center;
}

h1{
    color:#1e3a8a;
    margin-bottom:20px;
}

.progress{
    width:100%;
    height:12px;
    background:#d1d5db;
    border-radius:20px;
    overflow:hidden;
    margin-bottom:10px;
}

#progressBar{
    width:0%;
    height:100%;
    background:#2563eb;
    transition:.3s;
}

#progressText{
    margin-bottom:20px;
    color:#555;
}

.card{
    width:100%;
    height:380px;
    perspective:1000px;
    margin-bottom:20px;
}

.card-inner{
    position:relative;
    width:100%;
    height:100%;
    transition:transform .6s;
    transform-style:preserve-3d;
}

.card.flipped .card-inner{
    transform:rotateY(180deg);
}

.card-front,
.card-back{
    position:absolute;
    width:100%;
    height:100%;
    background:white;
    border-radius:20px;
    box-shadow:0 10px 25px rgba(0,0,0,.15);
    backface-visibility:hidden;
    display:flex;
    flex-direction:column;
    justify-content:center;
    align-items:center;
    padding:25px;
}

.card-back{
    transform:rotateY(180deg);
}

.hanzi{
    font-size:60px;
    font-weight:bold;
    color:#111827;
}

#pinyin{
    color:#2563eb;
    margin-bottom:10px;
}

#meaning{
    color:#16a34a;
    margin-bottom:20px;
}

.example{
    width:100%;
    text-align:left;
    line-height:1.8;
    font-size:18px;
}

.example hr{
    margin:15px 0;
}

button{
    border:none;
    border-radius:12px;
    padding:14px 22px;
    font-size:16px;
    cursor:pointer;
    transition:.2s;
}

button:hover{
    transform:translateY(-2px);
}

#flipBtn{
    width:100%;
    background:#2563eb;
    color:white;
    margin-bottom:15px;
}

.buttons{
    display:flex;
    gap:10px;
}

#knowBtn{
    flex:1;
    background:#22c55e;
    color:white;
}

#dontKnowBtn{
    flex:1;
    background:#ef4444;
    color:white;
}

@media(max-width:600px){

    .card{
        height:330px;
    }

    .hanzi{
        font-size:50px;
    }

    .example{
        font-size:16px;
    }

}
