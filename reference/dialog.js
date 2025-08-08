Office.initialize = () => {
  window.addEventListener("message", (event) => {
    const missing = JSON.parse(event.data);
    document.getElementById("message").innerText =
      "Missing elements: " + missing.join(", ");
  });
};

function send() {
  Office.context.ui.messageParent("send");
}

function cancel() {
  Office.context.ui.messageParent("cancel");
}
