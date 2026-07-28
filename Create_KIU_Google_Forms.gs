/**
 * Run createKiuPostgraduateAuditForms() once in Google Apps Script.
 * It creates separate PhD and Master's Google Forms plus linked response Sheets.
 * Copy the logged CONFIG_JS block into docs/config.js before publishing GitHub Pages.
 */
function createKiuPostgraduateAuditForms() {
  const folder = DriveApp.createFolder(
    "KIU Postgraduate Audit Forms " +
    Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd")
  );
  const phd = buildAuditForm_("PhD", true, folder);
  const masters = buildAuditForm_("Master's", false, folder);
  const result = {
    phdEditUrl: phd.form.getEditUrl(),
    phdPublishedUrl: phd.form.getPublishedUrl(),
    phdResponsesSheet: phd.sheet.getUrl(),
    mastersEditUrl: masters.form.getEditUrl(),
    mastersPublishedUrl: masters.form.getPublishedUrl(),
    mastersResponsesSheet: masters.sheet.getUrl(),
    folderUrl: folder.getUrl()
  };
  console.log(JSON.stringify(result, null, 2));
  console.log(
    "CONFIG_JS\\nwindow.KIU_AUDIT_FORMS = {\\n" +
    '  phd: "' + result.phdPublishedUrl + '",\\n' +
    '  masters: "' + result.mastersPublishedUrl + '"\\n' +
    "};"
  );
  return result;
}

function buildAuditForm_(level, isPhd, folder) {
  const form = FormApp.create(
    "KIU Western Campus - " + level + " Students' Audit and Progress-Monitoring Form"
  );
  form.setDescription(
    "Issued through the Research Coordinator in response to the DHDR postgraduate " +
    "students' audit directive dated 24 July 2026. Complete every applicable field " +
    "and submit privately. Do not post personal information in a WhatsApp group."
  );
  form.setConfirmationMessage(
    "Thank you. Your private audit response has been submitted to the Research Coordinator/DHDR team."
  );
  form.setAcceptingResponses(true);
  form.setProgressBar(true);
  form.setShuffleQuestions(false);
  form.setLimitOneResponsePerUser(false);

  addSection_(form, "Student and Programme Details");
  addText_(form, "Registration Number", true);
  addText_(form, "Full Name", true);
  addChoice_(form, "Sex", ["Female", "Male", "Prefer not to say", "Other"], true);
  addText_(form, "Telephone / WhatsApp Number", true);
  addText_(form, "Email Address", true);
  addText_(form, "School / Faculty", true);
  addText_(form, "Department", true);
  addText_(form, level + " Programme", true);
  addParagraph_(form, "Research Title" + (isPhd ? "" : " (if approved)"), isPhd);

  addSection_(form, "Registration and Academic Progress");
  addDate_(form, "Original Registration Date", true);
  addText_(form, "Academic Year of Registration (for example, 2024/2025)", true);
  addChoice_(form, "Current Registration Status",
    ["Active", "Inactive", "Deferred", "Completed", "Other"], true);
  const phdStages = [
    "Coursework", "Concept Development", "Concept Defence", "Proposal Development",
    "Proposal Defence", "Ethical Approval", "Data Collection", "Data Analysis",
    "Thesis Writing", "Pre-Viva Examination", "Final Viva", "Corrections",
    "Graduation Clearance", "Other"
  ];
  const mastersStages = [
    "Coursework", "Concept Development", "Concept Approval", "Proposal Development",
    "Proposal Defence", "Ethical Approval", "Data Collection", "Data Analysis",
    "Dissertation Writing", "Submission for Examination", "Viva Voce", "Corrections",
    "Graduation Clearance", "Other"
  ];
  addChoice_(form, "Current Academic Stage", isPhd ? phdStages : mastersStages, true);
  addDate_(form, "Date the Current Stage Began", true);
  addDate_(form, "Expected Date of Completing the Current Stage", true);

  addSection_(form, "Supervision");
  addText_(form, isPhd ? "Principal Supervisor's Full Name" : "Supervisor's Full Name", true);
  addText_(form, "Co-supervisor's Full Name (write N/A if not applicable)", false);
  addChoice_(form, "Are you in regular contact with your supervisor(s)?", ["Yes", "No"], true);
  addDate_(form, "Date of Last Supervisory Meeting", true);

  if (isPhd) {
    addSection_(form, "Doctoral Committee");
    addChoice_(form, "PhD Doctoral Committee Established", ["Yes", "No", "In Progress"], true);
    addText_(form, "Doctoral Committee Chairperson", false);
    addText_(form, "Committee Principal Supervisor", false);
    addText_(form, "Committee Co-supervisor", false);
    addText_(form, "Committee Member 1", false);
    addText_(form, "Committee Member 2", false);
    addText_(form, "External / Adjunct Member (if applicable)", false);
    addDate_(form, "Date of Last Doctoral Committee Meeting", false);
  }

  addSection_(form, "Milestones, Challenges and Support");
  addParagraph_(form, "Next Expected Academic Milestone", true);
  addDate_(form, "Target Date for the Next Milestone", true);
  addParagraph_(form, "Major Academic or Supervisory Challenge (write None if not applicable)", true);
  addParagraph_(form, "Support Required from DHDR", true);
  addText_(form, "Expected Year of Completion", true);
  addParagraph_(form, "Any Additional Information", false);

  addSection_(form, "Student Declaration");
  form.addCheckboxItem()
    .setTitle(
      "I confirm that the information provided is accurate and may be used by KIU Western " +
      "Campus for postgraduate academic monitoring, planning, official reminders, and appropriate support."
    )
    .setChoiceValues(["I confirm"])
    .setRequired(true);
  addText_(form, "Student's Full Name for Declaration", true);
  addDate_(form, "Date Submitted", true);

  const sheet = SpreadsheetApp.create(
    "RESPONSES - KIU " + level + " Students' Audit"
  );
  form.setDestination(FormApp.DestinationType.SPREADSHEET, sheet.getId());
  DriveApp.getFileById(form.getId()).moveTo(folder);
  DriveApp.getFileById(sheet.getId()).moveTo(folder);
  return { form: form, sheet: sheet };
}

function addSection_(form, title) {
  form.addPageBreakItem().setTitle(title);
}
function addText_(form, title, required) {
  form.addTextItem().setTitle(title).setRequired(required);
}
function addParagraph_(form, title, required) {
  form.addParagraphTextItem().setTitle(title).setRequired(required);
}
function addDate_(form, title, required) {
  form.addDateItem().setTitle(title).setIncludesYear(true).setRequired(required);
}
function addChoice_(form, title, values, required) {
  form.addMultipleChoiceItem().setTitle(title).setChoiceValues(values).setRequired(required);
}
