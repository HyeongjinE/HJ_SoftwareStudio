const timeStamp = document.querySelector(".main__timestamp");

const today = new Date();
const year = today.getFullYear();
const month = today.getMonth() + 1;
const date = today.getDate();
let day = today.getDay();

if (day==0) {
    day = "일"
} else if(day==1) {
    day = "월"
} else if(day==2) {
    day = "화"
} else if(day==3) {
    day = "수"
} else if(day==4) {
    day = "목"
} else if(day==5) {
    day = "금"
}else if(day==6) {
    day = "토"
}

timeStamp.innerText = year + "년 " + month + "월 " + date + "일(" + day + "요일)"

const develop = document.querySelector(".develop");

function developClick() {
    alert("과제를 주세요!")
}

develop.addEventListener("click", developClick)