Office.actions.associate("onMessageSendHandler", function (event) {
  event.completed({
    allowEvent: false,
    errorMessage: "Med-Trix TEST: Send event is working."
  });
});
