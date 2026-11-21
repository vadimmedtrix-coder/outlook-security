
"use strict";

Office.initialize = function () {
  console.log("Med-Trix: Office initialized");
};

const INTERNAL_DOMAINS = new Set(["med-trix.com"]);

const WARNING =
  "שים לב – שליחת מייל מחוץ לארגון\n\n" +
  "הודעה זו מיועדת לנמען אחד או יותר מחוץ לארגון.\n\n" +
  "לפני השליחה, יש לוודא שההודעה והקבצים המצורפים " +
  "אינם מכילים מידע רגיש או חסוי שאין לשתף מחוץ לארגון.\n\n" +
  "האם ברצונך להמשיך בשליחת המייל?";

const ERROR =
  "לא ניתן להשלים את בדיקת הנמענים. " +
  "יש לוודא שכתובות הנמענים תקינות לפני השליחה.";

function getRecipients(field) {
  return new Promise((resolve, reject) => {
    field.getAsync((result) => {
      if (
        result.status === Office.AsyncResultStatus.Succeeded &&
        Array.isArray(result.value)
      ) {
        resolve(result.value);
      } else {
        reject(new Error("Failed to retrieve recipients"));
      }
    });
  });
}

function isExternal(recipient) {
  const email = String(recipient.emailAddress || "")
    .trim()
    .toLowerCase();

  const atIndex = email.lastIndexOf("@");

  if (atIndex <= 0 || atIndex === email.length - 1) {
    return true;
  }

  const domain = email.substring(atIndex + 1);
  return !INTERNAL_DOMAINS.has(domain);
}

async function onMessageSendHandler(event) {
  let finished = false;

  function complete(options) {
    if (finished) return;
    finished = true;
    clearTimeout(timeout);
    event.completed(options);
  }

  const timeout = setTimeout(() => {
    complete({
      allowEvent: false,
      errorMessage: ERROR
    });
  }, 4000);

  try {
    const item = Office.context.mailbox.item;

    const [to, cc, bcc] = await Promise.all([
      getRecipients(item.to),
      getRecipients(item.cc),
      getRecipients(item.bcc)
    ]);

    const recipients = [...to, ...cc, ...bcc];

    if (recipients.some(isExternal)) {
      complete({
        allowEvent: false,
        errorMessage: WARNING
      });
    } else {
      complete({ allowEvent: true });
    }
  } catch (error) {
    console.error("Med-Trix recipient check failed", error);

    complete({
      allowEvent: false,
      errorMessage: ERROR
    });
  }
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);
