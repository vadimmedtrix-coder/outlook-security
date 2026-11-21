
"use strict";

// Med-Trix External Mail Security
// Outlook Smart Alerts - OnMessageSend

Office.initialize = function () {
  console.log("Med-Trix: Office initialized");
};

const INTERNAL_DOMAINS = new Set(["med-trix.com"]);
const RTL = "\u200F";

const ERROR_MESSAGE =
  RTL + "מערכות מידע – לא ניתן להשלים את בדיקת הנמענים. " +
  "יש לוודא שכתובות הדוא״ל תקינות לפני השליחה.";

function getRecipients(field) {
  return new Promise((resolve, reject) => {
    if (!field || typeof field.getAsync !== "function") {
      reject(new Error("Recipient field unavailable"));
      return;
    }

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

function getEmail(recipient) {
  return String(recipient?.emailAddress || "")
    .trim()
    .toLowerCase();
}

function isExternal(recipient) {
  const email = getEmail(recipient);
  const atIndex = email.lastIndexOf("@");

  if (atIndex <= 0 || atIndex === email.length - 1) {
    return true;
  }

  const domain = email.slice(atIndex + 1);

  return !INTERNAL_DOMAINS.has(domain);
}

function buildWarning(externalEmails) {
  const intro =
    RTL + "היי, כאן צוות מערכות מידע.\n\n" +
    RTL + "זוהתה הודעת דוא״ל המיועדת לנמען אחד או יותר " +
    "מחוץ לארגון.\n\n" +
    RTL + "להלן כתובות הדוא״ל החיצוניות:\n";

  const ending =
    "\n\n" +
    RTL + "לפני השליחה, חשוב לוודא שההודעה והקבצים " +
    "המצורפים אינם כוללים מידע רפואי, אישי, רגיש " +
    "או חסוי שאין לשתף מחוץ לארגון.\n\n" +
    RTL + "האם להמשיך בשליחת ההודעה?";

  const maxLength = 500;
  const lines = [];
  let shown = 0;

  for (let i = 0; i < externalEmails.length; i++) {
    const email = externalEmails[i];
    const line = `${shown + 1}. ${email}`;

    const remaining = externalEmails.length - shown - 1;

    const extra = remaining > 0
      ? "\n" + RTL + "ועוד " + remaining + " נמענים"
      : "";

    const candidate = [...lines, line].join("\n");

    if (
      (intro + candidate + extra + ending).length > maxLength
    ) {
      break;
    }

    lines.push(line);
    shown++;
  }

  if (shown === 0) {
    return (
      intro +
      RTL + "זוהו " + externalEmails.length +
      " נמענים חיצוניים. הרשימה ארוכה מדי להצגה." +
      ending
    ).slice(0, maxLength);
  }

  const remaining = externalEmails.length - shown;

  if (remaining > 0) {
    lines.push(RTL + "ועוד " + remaining + " נמענים");
  }

  return intro + lines.join("\n") + ending;
}

async function onMessageSendHandler(event) {
  let finished = false;
  let timeoutId;

  function complete(options) {
    if (finished) return;

    finished = true;
    clearTimeout(timeoutId);

    event.completed(options);
  }

  timeoutId = setTimeout(() => {
    console.error("Med-Trix: Recipient check timed out");

    complete({
      allowEvent: false,
      errorMessage: ERROR_MESSAGE
    });
  }, 4000);

  try {
    const item = Office.context.mailbox.item;

    const [to, cc, bcc] = await Promise.all([
      getRecipients(item.to),
      getRecipients(item.cc),
      getRecipients(item.bcc)
    ]);

    const allRecipients = [...to, ...cc, ...bcc];

    const externalEmails = [
      ...new Set(
        allRecipients
          .filter(isExternal)
          .map(getEmail)
      )
    ];

    if (externalEmails.length === 0) {
      complete({ allowEvent: true });
      return;
    }

    complete({
      allowEvent: false,
      errorMessage: buildWarning(externalEmails)
    });

  } catch (error) {
    console.error("Med-Trix: Send check failed", error);

    complete({
      allowEvent: false,
      errorMessage: ERROR_MESSAGE
    });
  }
}

Office.actions.associate(
  "onMessageSendHandler",
  onMessageSendHandler
);

console.log("Med-Trix: Handler registered");
