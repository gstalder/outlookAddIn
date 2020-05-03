'use strict'

let xhttp = new XMLHttpRequest();
xhttp.onreadystatechange = function () {
    if (this.readyState == 4 && this.status == 200) {
        document.getElementById("mostRecentJson").innerHTML =
            this.responseText;
    }
};
xhttp.open("GET", "most-recent-message.json", true);
xhttp.send();
