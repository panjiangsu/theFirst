//请对本示例加上注释；另外，完善和修正软件可能的逻辑错误。

(function start() {
    var img = document.querySelector("img");
    var heading = document.querySelector("h1");
    var myButton = document.querySelector('button');

    var author = document.querySelector("span");
    if (!localStorage.getItem('name')) {
        setUserName();
    }
    else {
        var storedName = localStorage.getItem('name');
        author.textContent = 'Mozilla is cool, ' + storedName;
    }

    myButton.onclick = function () {
        setUserName();
    }

    img.onclick = function () {
        var src = img.getAttribute("src");
        if (src === "images/田园风光.jpg") {
            var pass = prompt("Please input the password:");
            if (pass === "214431") {
                img.setAttribute("src", "images/另类风光.jpg");
                heading.innerHTML = "This is a special scenery!";
            }
        }
        else {
            img.setAttribute("src", "images/田园风光.jpg");
            heading.innerHTML = "这是一幅美丽的风景画";
        }
    }

    function setUserName() {
        var myName = prompt('Please enter your name:');

        if (!isEmptyName(myName)) {
            localStorage.setItem('name', myName);
            author.textContent = 'Mozilla is cool, ' + myName;
        }
        else {
            author.textContent = 'Mozilla is cool, ' + 'default user.';
        }

        function isEmptyName(name) {
            if (name === null || name === false) { return true; }
            if (name) {
                return (name === "") || (name.search(/^ *$/) != -1 ? true : false);
            }
            else {
                return false;
            }
        }

    }
}());