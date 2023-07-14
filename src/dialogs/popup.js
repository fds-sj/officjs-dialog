(async () => {
  await Office.onReady();

  document.getElementById("ok-button").onclick = sendCookieSetMessageToParent;

  function sendCookieSetMessageToParent() {
    const userName = document.getElementById("name-box").value;
    console.log(userName);
    setCookie('username', userName, 5);
    Office.context.ui.messageParent('cookie-set');
  }

  function setCookie(cname, cvalue, exdays) {
    const d = new Date();
    d.setTime(d.getTime() + (exdays*24*60*60*1000));
    let expires = "expires="+ d.toUTCString();
    document.cookie = cname + "=" + cvalue + ";" + expires + ";path=/" + ";SameSite=None; Secure";
    console.log(document.cookie);
  }
})();


