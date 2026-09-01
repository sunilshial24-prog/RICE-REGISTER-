STEP 5A — Sync Result Fix

Fixes the runtime error:
  SYNC ERROR: result is not defined

The Sync button now stores the returned result from fullSync() before reading result.count.

Also bumps the service-worker cache version so the corrected index.html is activated after deployment.

Google Apps Script URL remains the current configured endpoint.
