
"use strict";

Office.initialize = function () {
  console.log("Med-Trix: Office initialized");
};

function onMessageSendHandler(event) {
  console.log("Med-Trix: OnMessageSend triggered");

  event.completed({
    allowEvent: false,
    errorMessage: "Med-Trix TEST: Send event is working."
  });
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);

console.log("Med-Trix: Handler registered");
