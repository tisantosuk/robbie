# Robbie

Personal voice companion for Ti Santos (spike v3.1). Zero cost: GitHub Pages (this page), Google Apps Script backend in Ti's account, Gemini API free tier, phone's own voice. No billing account anywhere.

This repository holds no keys, passphrases, spreadsheet IDs or personal data. The backend address and passphrase reach the phone once, through the setup link, and stay in that phone's browser.

- `index.html`, `app.js`: phone page (voice detection on the phone, turn-taking on Ti's signal, tap to interrupt).
- `apps-script/Code.gs`: backend to paste into a new Apps Script project and deploy as a web app (Execute as: Me; Who has access: Anyone).
- `vad/`, `ort/`: voice detection model and runtime (from @ricky0123/vad-web 0.0.31 and onnxruntime-web 1.22.0, MIT licence).
