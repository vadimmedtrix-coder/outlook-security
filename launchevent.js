"use strict";

Office.initialize = function () {
  console.log("Med-Trix: Office initialized");
};

const INTERNAL_DOMAINS = new Set(["med-trix.com"]);
const RLM = "\u200F";

const WARNING = [
  `${RLM}היי, כאן מערכות מידע.`,
  `${RLM}זוהתה הודעת דוא"ל המיועדת לנמען אחד או יותר מחוץ לארגון.`,
  `${RLM}לפני השליחה, יש לוודא שההודעה והקבצים המצורפים אינם כוללים מידע רפואי, אישי, רגיש או חסוי שאין לשתף מחוץ לארגון.`,
  `${RLM}אם בדקת את תוכן ההודעה והקבצים המצורפים והשליחה מאושרת — ניתן להמשיך.`,
  `${RLM}האם ברצונך להמשיך בשליחת המייל?`
].join("\n\n");

const ERROR_TEXT = `${RLM}לא ניתן להשלים את בדיקת הנמענים. יש לבדוק את הנמענים לפני השליחה.`;

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

  const domain = email.slice(atIndex + 1);
  return !INTERNAL_DOMAINS.has(domain);
}

async function onMessageSendHandler(event) {
  try {
    const item = Office.context.mailbox.item;

    const allRecipients = (
      await Promise.all([
        getRecipients(item.to),
        getRecipients(item.cc),
        getRecipients(item.bcc)
      ])
    ).flat();

    const hasExternalRecipient = allRecipients.some(isExternal);

    if (hasExternalRecipient) {
      event.completed({
        allowEvent: false,
        errorMessage: WARNING
      });
    } else {
      event.completed({
        allowEvent: true
      });
    }
  } catch (error) {
    console.error("Med-Trix send check failed:", error);
    event.completed({
      allowEvent: false,
      errorMessage: ERROR_TEXT
    });
  }
}

Office.actions.associate("onMessageSendHandler", onMessageSendHandler);
