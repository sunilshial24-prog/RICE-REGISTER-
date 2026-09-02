STEP 6 — Major safety and usability update

Includes:
- Total pending stock across AAY + PHH + SFSS on Dashboard
- Real OFFLINE / CONNECTED / ONLINE-CHECK status
- Persistent IndexedDB schema upgrade without destructive migrations
- App-level and server-level data deletion disabled
- Backup and restore (merge only; never deletes existing records)
- Ledger Print / Save PDF (government-style A4 landscape)
- Ledger Excel export
- English / Hindi / Odia language selector
- Dealer/FPS settings
- Login: username P100, password 0201p100
- Apps Script URL remains hidden from Settings
- New service-worker cache version

Deploy:
1. Replace index.html, manifest.json, sw.js, icons.
2. Replace Apps Script with Code.gs.
3. Run setup().
4. Redeploy the Web App as a new version.
