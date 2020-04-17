var myHeadings = document.getElementsByTagName("h1");
for (var i=0;i<myHeadings.length;i++){
    myHeadings[i].addEventListener("click",headingAlert);
}

var myDocument = document.querySelector("html");
myDocument.addEventListener("click", documentAlert);

function headingAlert() {
    this.innerText = "A GOOD POEM";
}

function documentAlert() {
    window.alert("Don't click me again.")
}

