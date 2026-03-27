# MyNotes Outlook Add-in Starter

This starter is designed to work with the current MyNotes email intake flow.

## What it does
- Reads the currently opened Outlook email
- Builds `mynotes-email-capture` JSON
- Copies the payload to the clipboard
- Lets the user paste/import it into the local MyNotes app

## Files
- `manifest.xml` — Outlook add-in manifest
- `taskpane.html` — task pane UI
- `taskpane.js` — email extraction and JSON builder
- `commands.html` — required function file placeholder

## Important setup notes
This starter uses `https://localhost:3000/...` URLs in the manifest.
You will need to host these files locally over HTTPS for sideloading.

## Typical dev flow
1. Serve this folder locally over HTTPS.
2. Sideload `manifest.xml` into Outlook.
3. Open an email in Outlook.
4. Click **Send to MyNotes**.
5. Click **Copy MyNotes JSON**.
6. In MyNotes, use **Paste from clipboard** in Email Intake.

## Notes
- This is a starter scaffold, not a finished production add-in.
- Depending on Outlook version and host, some mailbox fields may vary slightly.
- The MyNotes app already supports auto-routing after import, so the add-in can stay simple.
