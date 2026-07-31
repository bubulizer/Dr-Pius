# KIU SEAS Postgraduate Audit

This package provides the no-charge deployment:

- GitHub Pages: public, mobile-friendly landing page and WhatsApp sharing.
- Google Forms: one private student form with programme-level and department routing.
- Google Sheets: a master response sheet owned by the Associate Dean Research.
- Protected department workbooks: student responses plus administrator-only doctoral-committee columns.

## Ownership and access

Run the deployment from `adresearch.seas@kiu.ac.ug`. The script gives department workbook access to:

- Civil Engineering: `rc.civil@kiu.ac.ug`
- Electrical, Telecommunications and Computer Engineering: `rc.etc@kiu.ac.ug`
- Mechanical Engineering: `bubupius.e@kiu.ac.ug`
- Biomedical Engineering: Associate Dean Research only until a coordinator email is added.

Coordinators can edit the doctoral-committee administration columns in their department workbook.
Student-submitted columns are protected. The public student form never displays doctoral-committee fields.

## Create the revised form

1. Sign into the Google account `adresearch.seas@kiu.ac.ug`.
2. Open <https://script.google.com/> and create a project.
3. Replace the default code with `Create_KIU_Google_Forms.gs`.
4. Run `createSeasPostgraduateAuditSystem`.
5. Approve the requested Forms, Sheets, Drive, and trigger permissions.
6. Open the created form, select **Published** > **Manage responders**, and change
   **Responder view** to **Anyone with the link**. This prevents students outside
   the KIU Google domain from being forced to sign in.
7. Copy the logged `CONFIG_JS` block into `docs/config.js`.
8. Submit one test response for each department and verify the routing.

Do not share form-edit, master-response, folder, or department-administration URLs publicly.
Only the published student-form URL belongs in `docs/config.js`.
