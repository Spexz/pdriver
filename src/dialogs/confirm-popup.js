(function () {
  "use strict";

  Office.onReady().then(function () {
    document.getElementById("yes-button").onclick = () => sendStringToParentPage("YES");
    document.getElementById("no-button").onclick = () => sendStringToParentPage("NO");
  });

  function sendStringToParentPage(selection) {
    // const userName = document.getElementById("name-box").value;
    Office.context.ui.messageParent(selection);
  }
})();
