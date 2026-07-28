# KIU Postgraduate Audit - GitHub Pages + Google Forms

This no-charge version uses:

- **GitHub Pages** for the public, WhatsApp-shareable landing page.
- **Google Forms** for private PhD and master's submissions.
- **Google Sheets** for responses visible only to the form owner's Google account.

## Create the Google Forms

1. Open `https://script.google.com/` using the Google account that should own the forms.
2. Create a **New project**.
3. Replace the default code with `Create_KIU_Google_Forms.gs`.
4. Run `createKiuPostgraduateAuditForms`.
5. Review and approve Google's requested Forms, Sheets, and Drive permissions.
6. Copy the `CONFIG_JS` block from the execution log into `docs/config.js`.

The script creates two forms, two linked response spreadsheets, and a Drive folder
containing all four files.

## Publish with GitHub Pages

Push this folder's contents to the `main` branch. The included workflow deploys
the `docs` directory. In the repository's **Settings â†’ Pages**, select **GitHub
Actions** if Pages is not enabled automatically.

Do not share the form edit URLs or response-spreadsheet URLs. Only the two published
form URLs belong in `docs/config.js` and WhatsApp messages.
