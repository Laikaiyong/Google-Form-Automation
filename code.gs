const APUBCC_OUTLOOK_MAIL = "apubcc@outlook.com";

const APUBCC_MEMBERSHIP_RESPONSES_COLUMN_EMAIL = "Email";
const APUBCC_MEMBERSHIP_RESPONSES_COLUMN_TP_NUMBER = "TP Number";
const APUBCC_MEMBERSHIP_RESPONSES_COLUMN_FIRST_NAME =
  "Full Name (as per IC)";


function setupTrigger() {
  ScriptApp.newTrigger('onFormSubmit')
    .forForm('1fHxIaCGpeEulE1NgoOtFn_xPiKZwYqLgC1y1FKSO_z0')
    .onFormSubmit()
    .create();
}

async function createEmailBody(applicantEmail, firstName) {
  const template = HtmlService.createTemplateFromFile("emailTemplate");

  template.NAME = firstName;
  const htmlMessage = template.evaluate().getContent();

  return htmlMessage;
}

async function onFormSubmit(e) {
  var itemResponses = await e.namedValues;
  console.log(itemResponses);

  var firstName = itemResponses[APUBCC_MEMBERSHIP_RESPONSES_COLUMN_FIRST_NAME][0];
  console.log(firstName);

  var email = itemResponses[APUBCC_MEMBERSHIP_RESPONSES_COLUMN_EMAIL][0];
  console.log(email);

  var tpNumber = itemResponses[APUBCC_MEMBERSHIP_RESPONSES_COLUMN_TP_NUMBER][0];
  console.log(tpNumber);

  var emailSubject = `Welcome to APU BCC Alpha Access `;
  var message = await createEmailBody(email, firstName);

  await GmailApp.sendEmail(email, emailSubject, message, {
    from: APUBCC_OUTLOOK_MAIL,
    cc: `${tpNumber}@mail.apu.edu.my`,
    name: "APU Blockchain & Cryptocurrency Club",
    htmlBody: message,
  });
}
