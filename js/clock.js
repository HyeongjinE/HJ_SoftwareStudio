const clock = document.querySelector(".main__timestamp span:last-child");

function getClock() {
    const date = new Date();
    const hours = String(date.getHours()).padStart(2,"0");
    const minutes = String(date.getMinutes()).padStart(2,"0");
    const seconds = String(date.getSeconds()).padStart(2,"0");
    clock.innerText = `${hours}:${minutes}:${seconds}`;

    /*
    const date2 = new Date("2023-12-25 GMT+0900");

    const dDay = date2 - date;
    const day = Math.floor(dDay/(1000*60*60*24));
    const hour = Math.floor((dDay/(1000*60*60)) % 24);
    const minute = Math.floor((dDay/(1000*60)) % 60);
    const second = Math.floor((dDay/1000) % 60);
    clock2.innerText = `${day}d ${hour}h ${minute}m ${second}s`;
    */
}

getClock();
setInterval(getClock, 1000);
