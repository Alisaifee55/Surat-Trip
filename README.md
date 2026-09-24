# SplitEase — 4 Person

Expense splitter for Ali Asgar, Hakim, Taher, and Yusuf. Same visual style
as Ali's SplitEase, backed by its own separate Google Sheet.

## What's already wired up
- `Code.gs` writes explicitly to this Sheet: https://docs.google.com/spreadsheets/d/11go52W0cRlwYDnPlJ6XVDk7WFr2Z8pFGRhIJmL5_jf8/edit
- `index.html` calls this Apps Script deployment: https://script.google.com/macros/s/AKfycbxVhDSD_01xu4o7S-XBQkDiJxyHmuSrNpqw9Y0laH56aDZ3CkcGK_lWkI2MydIju-nbJA/exec
- People are set to: Ali Asgar, Hakim, Taher, Yusuf (edit the `PEOPLE` array at the
  top of `Code.gs` if this ever changes — no other code needs to change)

## One-time backend setup (Google Apps Script)
1. Open the Google Sheet linked above.
2. Extensions → Apps Script.
3. Delete any existing code, paste in `Code.gs` from this folder.
4. Deploy → New deployment → Web app → Execute as: Me → Who has access: Anyone.
5. If you ever create a *new* deployment (new URL), update `EXEC_URL` near the
   top of `index.html`'s `<script>` block to match, and re-deploy the frontend.

## Deploying the frontend

### Option A — GitHub + Cloudflare Pages (recommended, matches what you asked for)
1. Create a new GitHub repository and push every file in this folder to it
   (keep the folder structure exactly as-is — `icons/` must stay a subfolder).
2. In Cloudflare dashboard → Workers & Pages → Create → Pages → Connect to Git
   → select this repository.
3. Build settings: **Framework preset: None**, **Build command: (leave blank)**,
   **Build output directory: /** (project root).
4. Deploy. Cloudflare will give you a `*.pages.dev` URL.

### Option B — Direct upload to Cloudflare Pages (no GitHub needed)
1. Cloudflare dashboard → Workers & Pages → Create → Pages → Upload assets.
2. Drag this entire folder in and deploy.

## Files in this folder
- `index.html` — the whole app (self-contained, logo embedded inline)
- `Code.gs` — Google Apps Script backend
- `manifest.json` — enables "Add to Home Screen" / PWA install with the logo
- `_headers` — Cloudflare Pages caching rules (icons cached hard, index.html always fresh)
- `icons/` — favicon, Apple touch icon, Android/PWA icons (same logo as Ali's SplitEase)
