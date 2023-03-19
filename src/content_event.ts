(function () {
  "use strict";
  const messageName = "teams-see-more-message";
  chrome.runtime.onMessage.addListener((request, sender, sendResponse) => {
    if (request.message === messageName) {
      document.querySelectorAll(".ts-see-more-fold").forEach((node) => {
        // console.log(node);
        node.dispatchEvent(new window.MouseEvent("click", { bubbles: true }));
      });
      // document.querySelectorAll('.ts-collapsed-common').forEach(node => {
      //     // console.log(node);
      //     node.dispatchEvent(new window.MouseEvent("click", { bubbles: true }));
      // });
      sendResponse({ message: "success" });
      return true;
    }
  });
})();
