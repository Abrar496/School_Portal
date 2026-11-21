# School Server & GitHub Portal v2 Setup

This bundle upgrades your School Server (`Code.gs`) and GitHub Pages Student Portal (`web/index.html`) so students and guardians can view:
- **Notices & Homework** (with Class Captain posting)
- **Attendance & Fine**
- **Every Subject's Marks** (published directly from **Abrar EduHub → Notes → Marks**)
- **Full Student Profile Information** (synced directly from **Abrar EduHub → Classes → Students**)

## Quick 2-Step Update
1. **Update Google Apps Script (`Code.gs`)**:
   - Open your Google Sheet → **Extensions → Apps Script**.
   - Replace the code in `Code.gs` with the new `Code.gs` from this ZIP and click **Save**.
   - Click **Deploy → Manage deployments → Edit (pencil icon) → Version: New version → Deploy** (this keeps your exact same `/exec` URL!).
   - Run `setup()` once from the Apps Script toolbar to create the `Student_Profiles` sheet.
2. **Update GitHub Pages (`web/index.html`)**:
   - Open your `abrar496/School_Portal` repository on GitHub and replace `index.html` with `web/index.html` from this ZIP.
   - Commit the change — GitHub Pages updates automatically in ~1 minute.
