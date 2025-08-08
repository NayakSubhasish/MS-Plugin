function validateEmail(event) {
  const item = Office.context.mailbox.item;

  item.body.getAsync("text", (result) => {
    const bodyText = result.value;
    const missing = [];

    if (!bodyText.includes("Dear")) missing.push("Greeting");
    if (!bodyText.includes("Regards")) missing.push("Closing");
    if (!bodyText.match(/Please|Action Required/)) missing.push("Action Phrase");

    if (missing.length > 0) {
      Office.context.ui.displayDialogAsync(
        "https://yourdomain.com/dialog.html",
        { height: 40, width: 30 },
        (asyncResult) => {
          const dialog = asyncResult.value;
          dialog.messageChild(JSON.stringify(missing));

          dialog.addEventHandler(Office.EventType.DialogMessageReceived, (arg) => {
            if (arg.message === "send") {
              event.completed({ allowEvent: true });
            } else {
              event.completed({ allowEvent: false });
            }
            dialog.close();
          });
        }
      );
    } else {
      event.completed({ allowEvent: true });
    }
  });
}
