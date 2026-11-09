# CHARACTER ARCHIVE · Cloudflare edition

Frontend: Cloudflare Pages (`public/`)
Backend: Google Apps Script (`gas/Code.gs`)
Database: Google Sheets + Google Drive

## 1. Apps Script backend
1. Open the existing Apps Script project that owns the Character Archive spreadsheet.
2. Replace `Code.gs` with `gas/Code.gs`. Keep the existing `Index.html` file if present; Cloudflare no longer serves it.
3. Save and run `setup()` once. Approve the Google Sheets and Drive permissions.
4. Open the execution log and copy the value after `Cloudflare GAS_API_SECRET`. Keep it private.
5. Deploy > New deployment > Web app. Execute as **Me**; access **Anyone**. This is required because Cloudflare calls the endpoint. The secret is checked server-side and is never sent to the browser.
6. Copy the deployed `/exec` URL.

## 2. GitHub / Cloudflare Pages
1. Push this repository's contents to the GitHub repository root.
2. In Cloudflare Pages, connect the repository.
3. Framework preset: **None**. Build command: leave blank / None. Build output directory: `public`.
4. In Pages > Settings > Variables and Secrets, add these runtime variables:
   - `GAS_WEB_APP_URL`: the Apps Script deployed `/exec` URL
   - `GAS_API_SECRET`: the secret copied from the Apps Script execution log (mark as secret)
5. Redeploy the latest commit.

## Notes
- The browser talks only to `/api/*` on the Pages domain. The Pages Function forwards approved actions to Apps Script, keeping the API secret out of frontend files.
- The app's existing interface is retained via a small `google.script.run` compatibility bridge.
- Public character links use `https://YOUR-PAGES-DOMAIN/?character=CHARACTER_ID`.
- This is a single-owner archive. Do not expose the Apps Script URL and secret together. Rotate the secret by changing the `CHARACTER_ARCHIVE_API_SECRET` script property and updating the Cloudflare secret.
