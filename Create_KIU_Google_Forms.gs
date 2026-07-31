/**
 * SEAS Postgraduate Audit - Google Forms/Sheets deployment script.
 *
 * IMPORTANT:
 * 1. Run this script while signed in as adresearch.seas@kiu.ac.ug.
 * 2. Run createSeasPostgraduateAuditSystem() once.
 * 3. Copy the logged CONFIG_JS block into docs/config.js.
 *
 * The public student form contains no doctoral-committee fields. Each response is
 * copied into a protected department administration workbook. Only the Associate
 * Dean Research and that department's coordinator receive access to the workbook.
 */

const SEAS_AUDIT = {
  ownerEmail: "adresearch.seas@kiu.ac.ug",
  deadline: "To be communicated by the Research Coordinator",
  departments: {
    "Civil Engineering": {
      key: "civil",
      coordinator: "rc.civil@kiu.ac.ug",
      programmes: [
        "Geotechnical Engineering",
        "Construction Management",
        "Environmental Engineering",
        "Water Resources Engineering",
        "Structural Engineering",
        "Transportation Engineering",
        "Other approved Civil Engineering specialisation"
      ]
    },
    "Electrical, Telecommunications and Computer Engineering": {
      key: "etc",
      coordinator: "rc.etc@kiu.ac.ug",
      programmes: [
        "Power Systems Engineering",
        "Telecommunications Engineering",
        "Computer Engineering",
        "Renewable Energy Engineering",
        "Other approved ETC Engineering specialisation"
      ]
    },
    "Mechanical Engineering": {
      key: "mechanical",
      coordinator: "bubupius.e@kiu.ac.ug",
      programmes: [
        "General Mechanical Engineering",
        "Industrial Engineering",
        "Renewable Energy Engineering",
        "Other approved Mechanical Engineering specialisation"
      ]
    },
    "Biomedical Engineering": {
      key: "biomedical",
      coordinator: "",
      programmes: [
        "Biomedical Engineering",
        "Other approved Biomedical Engineering specialisation"
      ]
    }
  }
};

const STUDENT_EXPORT_HEADERS = [
  "Submission Timestamp",
  "Registration Number",
  "Full Name",
  "Sex",
  "Telephone / WhatsApp Number",
  "Email Address",
  "School / Faculty",
  "PG Programme Level",
  "Department",
  "Programme / Specialisation",
  "Other Programme / Specialisation",
  "Research Title",
  "Original Registration Date",
  "Academic Year of Registration",
  "Current Registration Status",
  "Other Registration Status",
  "Current Academic Stage",
  "Other Current Academic Stage",
  "Date Current Stage Began",
  "Expected Date of Completing Current Stage",
  "Principal / Sole Supervisor",
  "Co-supervisor 1",
  "Co-supervisor 2",
  "Co-supervisor 3",
  "Co-supervisor 4",
  "Regular Contact with Supervisor(s)",
  "Date of Last Supervisory Meeting",
  "Next Expected Academic Milestone",
  "Target Date for Next Milestone",
  "Major Academic or Supervisory Challenge",
  "Support Required from SEAS Research Office / DHDR",
  "Expected Year of Completion",
  "Additional Information",
  "Student Declaration Name",
  "Date Submitted"
];

const COMMITTEE_ADMIN_HEADERS = [
  "Doctoral Committee Status",
  "Committee Chairperson",
  "Committee Principal Supervisor",
  "Committee Co-supervisor 1",
  "Committee Co-supervisor 2",
  "Committee Co-supervisor 3",
  "Committee Co-supervisor 4",
  "Committee Member 1",
  "Committee Member 2",
  "External / Adjunct Member",
  "Date of Last Doctoral Committee Meeting",
  "Administrative Notes",
  "Reviewed / Updated By",
  "Reviewed / Updated Date"
];

function createSeasPostgraduateAuditSystem() {
  verifyOwnerAccount_();
  const dateTag = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd");
  const folder = DriveApp.createFolder("SEAS Postgraduate Audit " + dateTag);
  const departmentFiles = createDepartmentAdministrationFiles_(folder);
  const studentForm = createStudentAuditForm_(folder);
  const masterSheet = createMasterResponseSheet_(studentForm, folder);

  const properties = {
    masterSpreadsheetId: masterSheet.getId(),
    formId: studentForm.getId()
  };
  Object.keys(departmentFiles).forEach(function(key) {
    properties["departmentSheet_" + key] = departmentFiles[key].spreadsheet.getId();
  });
  PropertiesService.getScriptProperties().setProperties(properties, true);

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction() === "routeAuditSubmissionToDepartment_") {
      ScriptApp.deleteTrigger(trigger);
    }
  });
  ScriptApp.newTrigger("routeAuditSubmissionToDepartment_")
    .forForm(studentForm)
    .onFormSubmit()
    .create();

  const result = {
    studentFormEditUrl: studentForm.getEditUrl(),
    studentFormPublishedUrl: studentForm.getPublishedUrl(),
    masterResponsesSheet: masterSheet.getUrl(),
    folderUrl: folder.getUrl(),
    departmentAdministration: {}
  };
  Object.keys(departmentFiles).forEach(function(key) {
    result.departmentAdministration[key] = departmentFiles[key].spreadsheet.getUrl();
  });

  console.log(JSON.stringify(result, null, 2));
  console.log(
    "CONFIG_JS\\nwindow.KIU_SEAS_AUDIT = {\\n" +
    '  studentForm: "' + result.studentFormPublishedUrl + '",\\n' +
    '  deadline: "' + SEAS_AUDIT.deadline.replace(/"/g, '\\"') + '"\\n' +
    "};"
  );
  return result;
}

function createStudentAuditForm_(folder) {
  const form = FormApp.create(
    "SEAS Postgraduate Students' Audit and Progress-Monitoring Form"
  );
  form.setDescription(
    "KAMPALA INTERNATIONAL UNIVERSITY, WESTERN CAMPUS\\n" +
    "SCHOOL OF ENGINEERING AND APPLIED SCIENCES | HIGHER DEGREES AND RESEARCH\\n\\n" +
    "Issued through the Associate Dean Research via the departmental Research Coordinator " +
    "in response to the DHDR postgraduate students' audit directive dated 24 July 2026. " +
    "The SEAS Research Office, working with DHDR, is conducting an audit of registered " +
    "postgraduate students to update official records, monitor academic progress, identify " +
    "supervisory gaps, and provide timely milestone support.\\n\\n" +
    "Complete every applicable field and submit privately by " + SEAS_AUDIT.deadline + ". " +
    "Do not post personal information in a WhatsApp group."
  );
  form.setConfirmationMessage(
    "Thank you. Your private SEAS postgraduate audit response has been submitted. " +
    "Contact your departmental Research Coordinator if you need to correct a response."
  );
  form.setAcceptingResponses(true);
  form.setProgressBar(true);
  form.setShuffleQuestions(false);
  form.setLimitOneResponsePerUser(false);

  addSection_(form, "1. Student and Programme Details",
    "Use the details shown in your official university records.");
  addText_(form, "Registration Number", true);
  addText_(form, "Full Name", true, "As shown in university records");
  addChoice_(form, "Sex", ["Female", "Male", "Prefer not to say", "Other"], true);
  addText_(form, "Telephone / WhatsApp Number", true, "Include country code where applicable");
  addText_(form, "Email Address", true);
  addText_(form, "School / Faculty", true, "School of Engineering and Applied Sciences");
  addList_(form, "PG Programme Level", ["MSc", "PhD"], true);

  const departmentItem = form.addMultipleChoiceItem()
    .setTitle("Department")
    .setHelpText("Choose your academic department. The next page will show its programme options.")
    .setRequired(true);

  const departmentSections = {};
  Object.keys(SEAS_AUDIT.departments).forEach(function(departmentName) {
    const settings = SEAS_AUDIT.departments[departmentName];
    const section = form.addPageBreakItem()
      .setTitle(departmentName + " Programme")
      .setHelpText("Choose the programme or specialisation applicable to your MSc or PhD study.");
    departmentSections[departmentName] = section;
    addList_(form, departmentName + " Programme / Specialisation", settings.programmes, true);
    addText_(form, departmentName + " - Other Programme / Specialisation", false,
      "Complete only if you selected an Other approved specialisation option");
  });

  const progressSection = form.addPageBreakItem()
    .setTitle("2. Registration and Academic Progress")
    .setHelpText("Provide your current registration and milestone status.");

  Object.keys(departmentSections).forEach(function(departmentName) {
    departmentSections[departmentName].setGoToPage(progressSection);
  });
  departmentItem.setChoices(
    Object.keys(departmentSections).map(function(departmentName) {
      return departmentItem.createChoice(departmentName, departmentSections[departmentName]);
    })
  );

  addParagraph_(form, "Research Title", true, "Current approved or working title");
  addDate_(form, "Original Registration Date", true);
  addText_(form, "Academic Year of Registration", true, "Example: 2024/2025");
  addChoice_(form, "Current Registration Status",
    ["Active", "Inactive", "Deferred", "Completed"], true, true);
  addText_(form, "Other Registration Status", false,
    "Complete only if you selected Other");
  addChoice_(form, "Current Academic Stage", [
    "Coursework",
    "Concept Development",
    "Concept Defence",
    "Proposal Development",
    "Proposal Defence",
    "REC",
    "Instruments Validation",
    "WIP 1",
    "WIP 2",
    "WIP 3",
    "WIP 4",
    "Integrated Results",
    "Final Preparation",
    "Internal Viva",
    "External Viva",
    "Corrections",
    "Graduation Clearance"
  ], true, true);
  addText_(form, "Other Current Academic Stage", false,
    "Complete only if you selected Other");
  addDate_(form, "Date Current Stage Began", true);
  addDate_(form, "Expected Date of Completing Current Stage", true);

  addSection_(form, "3. Supervision",
    "Provide the names of your current appointed supervisors.");
  addText_(form, "Principal / Sole Supervisor", true);
  addText_(form, "Co-supervisor 1", false, "Leave blank if not applicable");
  addText_(form, "Co-supervisor 2", false, "Leave blank if not applicable");
  addText_(form, "Co-supervisor 3", false, "Leave blank if not applicable");
  addText_(form, "Co-supervisor 4", false, "Leave blank if not applicable");
  addChoice_(form, "Regular Contact with Supervisor(s)", ["Yes", "No"], true);
  addDate_(form, "Date of Last Supervisory Meeting", true);

  addSection_(form, "4. Milestones, Challenges and Support",
    "Explain any support need clearly enough for appropriate follow-up.");
  addParagraph_(form, "Next Expected Academic Milestone", true);
  addDate_(form, "Target Date for Next Milestone", true);
  addParagraph_(form, "Major Academic or Supervisory Challenge", true,
    "Write None if not applicable");
  addParagraph_(form, "Support Required from SEAS Research Office / DHDR", true,
    "Be specific where possible");
  addText_(form, "Expected Year of Completion", true, "YYYY");
  addParagraph_(form, "Additional Information", false, "Optional");

  addSection_(form, "5. Student Declaration",
    "Your submission must relate only to your own academic record.");
  form.addCheckboxItem()
    .setTitle(
      "I confirm that the information provided is accurate and may be used by Kampala " +
      "International University - Western Campus for postgraduate academic monitoring, " +
      "planning, official reminders, and appropriate student support. I understand that " +
      "I should submit privately and must not share another student's information."
    )
    .setChoiceValues(["I confirm"])
    .setRequired(true);
  addText_(form, "Student Declaration Name", true, "Student's full name");
  addDate_(form, "Date Submitted", true);

  DriveApp.getFileById(form.getId()).moveTo(folder);
  return form;
}

function createMasterResponseSheet_(form, folder) {
  const spreadsheet = SpreadsheetApp.create(
    "MASTER RESPONSES - SEAS Postgraduate Students' Audit"
  );
  form.setDestination(FormApp.DestinationType.SPREADSHEET, spreadsheet.getId());
  DriveApp.getFileById(spreadsheet.getId()).moveTo(folder);
  return spreadsheet;
}

function createDepartmentAdministrationFiles_(folder) {
  const files = {};
  Object.keys(SEAS_AUDIT.departments).forEach(function(departmentName) {
    const settings = SEAS_AUDIT.departments[departmentName];
    const spreadsheet = SpreadsheetApp.create(
      "ADMIN - " + departmentName + " Postgraduate Audit"
    );
    const sheet = spreadsheet.getSheets()[0];
    sheet.setName("Student Audit and Committee Admin");
    const headers = STUDENT_EXPORT_HEADERS.concat(COMMITTEE_ADMIN_HEADERS);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, STUDENT_EXPORT_HEADERS.length)
      .setBackground("#087b2c")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setWrap(true);
    sheet.getRange(1, STUDENT_EXPORT_HEADERS.length + 1, 1, COMMITTEE_ADMIN_HEADERS.length)
      .setBackground("#c89a2b")
      .setFontColor("#ffffff")
      .setFontWeight("bold")
      .setWrap(true);
    sheet.autoResizeColumns(1, headers.length);
    sheet.setColumnWidths(1, headers.length, 150);

    protectStudentResponseColumns_(sheet);
    const file = DriveApp.getFileById(spreadsheet.getId());
    file.moveTo(folder);
    if (settings.coordinator) {
      file.addEditor(settings.coordinator);
    }
    files[settings.key] = { spreadsheet: spreadsheet, sheet: sheet };
  });
  return files;
}

function protectStudentResponseColumns_(sheet) {
  const range = sheet.getRange(
    1, 1, sheet.getMaxRows(), STUDENT_EXPORT_HEADERS.length
  );
  const protection = range.protect()
    .setDescription("Student-submitted fields - maintained by the Associate Dean Research");
  protection.setWarningOnly(false);
  const editors = protection.getEditors().filter(function(user) {
    return String(user.getEmail() || "").toLowerCase() !== SEAS_AUDIT.ownerEmail.toLowerCase();
  });
  if (editors.length) protection.removeEditors(editors);
  protection.addEditor(SEAS_AUDIT.ownerEmail);
  if (protection.canDomainEdit()) protection.setDomainEdit(false);
}

function routeAuditSubmissionToDepartment_(event) {
  if (!event || !event.response) {
    throw new Error("This function must run from the installed Google Form submit trigger.");
  }
  const responseMap = {};
  event.response.getItemResponses().forEach(function(itemResponse) {
    const title = itemResponse.getItem().getTitle();
    const value = itemResponse.getResponse();
    responseMap[title] = Array.isArray(value) ? value.join("; ") : value;
  });

  const department = String(responseMap["Department"] || "").trim();
  const settings = SEAS_AUDIT.departments[department];
  if (!settings) {
    console.error("Unrecognized department: " + department);
    return;
  }
  const spreadsheetId = PropertiesService.getScriptProperties()
    .getProperty("departmentSheet_" + settings.key);
  if (!spreadsheetId) throw new Error("Department administration spreadsheet is not configured.");

  const specialisation = firstNonEmpty_([
    responseMap[department + " Programme / Specialisation"],
    responseMap["Civil Engineering Programme / Specialisation"],
    responseMap["Electrical, Telecommunications and Computer Engineering Programme / Specialisation"],
    responseMap["Mechanical Engineering Programme / Specialisation"],
    responseMap["Biomedical Engineering Programme / Specialisation"]
  ]);
  const otherSpecialisation = firstNonEmpty_([
    responseMap[department + " - Other Programme / Specialisation"],
    responseMap["Civil Engineering - Other Programme / Specialisation"],
    responseMap["Electrical, Telecommunications and Computer Engineering - Other Programme / Specialisation"],
    responseMap["Mechanical Engineering - Other Programme / Specialisation"],
    responseMap["Biomedical Engineering - Other Programme / Specialisation"]
  ]);

  const rowMap = {
    "Submission Timestamp": event.response.getTimestamp(),
    "Registration Number": responseMap["Registration Number"],
    "Full Name": responseMap["Full Name"],
    "Sex": responseMap["Sex"],
    "Telephone / WhatsApp Number": responseMap["Telephone / WhatsApp Number"],
    "Email Address": responseMap["Email Address"],
    "PG Programme Level": responseMap["PG Programme Level"],
    "Department": department,
    "Programme / Specialisation": specialisation,
    "Other Programme / Specialisation": otherSpecialisation,
    "Research Title": responseMap["Research Title"],
    "Original Registration Date": responseMap["Original Registration Date"],
    "Academic Year of Registration": responseMap["Academic Year of Registration"],
    "Current Registration Status": responseMap["Current Registration Status"],
    "Other Registration Status": responseMap["Other Registration Status"],
    "Current Academic Stage": responseMap["Current Academic Stage"],
    "Other Current Academic Stage": responseMap["Other Current Academic Stage"],
    "Date Current Stage Began": responseMap["Date Current Stage Began"],
    "Expected Date of Completing Current Stage": responseMap["Expected Date of Completing Current Stage"],
    "Principal / Sole Supervisor": responseMap["Principal / Sole Supervisor"],
    "Co-supervisor 1": responseMap["Co-supervisor 1"],
    "Co-supervisor 2": responseMap["Co-supervisor 2"],
    "Co-supervisor 3": responseMap["Co-supervisor 3"],
    "Co-supervisor 4": responseMap["Co-supervisor 4"],
    "Regular Contact with Supervisor(s)": responseMap["Regular Contact with Supervisor(s)"],
    "Date of Last Supervisory Meeting": responseMap["Date of Last Supervisory Meeting"],
    "Next Expected Academic Milestone": responseMap["Next Expected Academic Milestone"],
    "Target Date for Next Milestone": responseMap["Target Date for Next Milestone"],
    "Major Academic or Supervisory Challenge": responseMap["Major Academic or Supervisory Challenge"],
    "Support Required from SEAS Research Office / DHDR": responseMap["Support Required from SEAS Research Office / DHDR"],
    "Expected Year of Completion": responseMap["Expected Year of Completion"],
    "Additional Information": responseMap["Additional Information"],
    "Student Declaration Name": responseMap["Student Declaration Name"],
    "Date Submitted": responseMap["Date Submitted"]
  };
  const row = STUDENT_EXPORT_HEADERS.map(function(header) {
    return rowMap[header] === undefined ? "" : rowMap[header];
  }).concat(COMMITTEE_ADMIN_HEADERS.map(function() { return ""; }));

  const sheet = SpreadsheetApp.openById(spreadsheetId)
    .getSheetByName("Student Audit and Committee Admin");
  sheet.appendRow(row);
}

function verifyOwnerAccount_() {
  const effectiveEmail = String(Session.getEffectiveUser().getEmail() || "").toLowerCase();
  if (effectiveEmail && effectiveEmail !== SEAS_AUDIT.ownerEmail.toLowerCase()) {
    throw new Error(
      "Run this deployment while signed in as " + SEAS_AUDIT.ownerEmail +
      " so the Associate Dean Research owns the form and administration files."
    );
  }
}

function firstNonEmpty_(values) {
  for (let index = 0; index < values.length; index += 1) {
    const value = values[index];
    if (value !== null && value !== undefined && String(value).trim() !== "") return value;
  }
  return "";
}

function addSection_(form, title, helpText) {
  const item = form.addPageBreakItem().setTitle(title);
  if (helpText) item.setHelpText(helpText);
  return item;
}

function addText_(form, title, required, helpText) {
  const item = form.addTextItem().setTitle(title).setRequired(required);
  if (helpText) item.setHelpText(helpText);
  return item;
}

function addParagraph_(form, title, required, helpText) {
  const item = form.addParagraphTextItem().setTitle(title).setRequired(required);
  if (helpText) item.setHelpText(helpText);
  return item;
}

function addDate_(form, title, required) {
  return form.addDateItem().setTitle(title).setIncludesYear(true).setRequired(required);
}

function addChoice_(form, title, values, required, allowOther) {
  const item = form.addMultipleChoiceItem()
    .setTitle(title)
    .setChoiceValues(values)
    .setRequired(required);
  if (allowOther) item.showOtherOption(true);
  return item;
}

function addList_(form, title, values, required) {
  return form.addListItem()
    .setTitle(title)
    .setChoiceValues(values)
    .setRequired(required);
}
