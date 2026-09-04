STEP 8 FINAL — Shared cloud sync

Google Apps Script:
1. Replace Code.gs.
2. Run setup() once.
3. Deploy a NEW Web App version.
4. Execute as Me; Who has access Anyone.
5. Keep the same /exec URL.

GitHub:
Replace index.html, manifest.json, sw.js, icon-192.png, icon-512.png.

AppMint:
Use the SAME GitHub Pages URL. Do not create a separate data store.

The app:
- Pushes pending entries to Google Sheets.
- Pulls cloud entries on startup, focus, visibility return, every 30 seconds, and when internet returns.
- Keeps pending local changes safe.
- Uses GET upsert for WebView compatibility and POST as fallback.
- Does not delete records.
