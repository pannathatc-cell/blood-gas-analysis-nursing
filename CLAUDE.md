# ABG NSSKRU — maintenance notes

Owner: อาจารย์ปัณณทัต จันพานิช, Faculty of Nursing, SSKRU. Claude maintains this site on the owner's behalf: when the owner asks for a change, edit, commit and push to `main` directly — GitHub Pages redeploys automatically (1–2 min).

## Live
- Site: https://pannathatc-cell.github.io/blood-gas-analysis-nursing/
- Pages source: branch `main`, folder `/` (root)
- Google Sheet: `1bV3ToecEFKpfnIgWJASQaVAMfgItPQ0hB_uPvSCUJnQ` (tabs `Results`, `Visits`)
- Apps Script web app (`SHEET_API` in index.html):
  https://script.google.com/macros/s/AKfycbzqc5HKnAx4ZGGd_cy7kfF_2PEgltkOupj-O2vfEIpaZXwMnD24u9z4Ix07yEcpqU-G/exec
  - `GET ?action=ping` → health check
  - `GET ?action=list&key=<admin code>` → results + visits (admin only)
  - `POST` text body JSON `{action: saveResult|saveVisit|batch, ...}`

## Files
- `index.html` — whole app in one file (CSS + JS + base64 WebP character images). UI language: Thai.
- `apps-script/Code.gs` — copy of the Apps Script. Editing it here does NOT deploy it: the owner must paste it into the Apps Script editor and use Manage deployments → Edit → New version (URL stays the same).
- `favicon.png`, `.nojekyll`, `README.md`.

## App structure (index.html)
- Tabs: `#interp` (ABG interpretation + nursing care + SBAR), `#game` (GBL: balloon / shooter / clinical), `#dash` and `#usage` (admin only).
- Players sign in with a nickname only. Admin = typing the admin code in the nickname box (checked against `ADMIN_HASH`, SHA-256). Admin key is kept in localStorage `abg.adminKey` and sent to Apps Script to read data.
- `Session` = time-on-app tracker: counts only visible + active (input within 5 min) time; heartbeat every 60 s and `sendBeacon` on hide/leave; posts `saveSession` → Visits cols E (ใช้งานล่าสุด) and F (เวลาใช้งาน (นาที)).
- `Store` = Google Sheets layer: saves to localStorage first, posts to `SHEET_API`, queues offline posts in `abg.queue`; admin dashboards pull every 30 s.
- Interpretation: `interpretABG(ph, paco2, hco3, o2rx, pao2, be)`; normal PaO₂ 80–100 mmHg, BE −2…+2 mmol/L. SpO₂ was removed on request.
- Characters: `MASCOT` map of data URIs; slots use `data-mascot="<key>"`, popup text via `data-say`. Male nurse salute frames (`boyN_*`) animate a wink via `img.frames`.

## Checks before pushing
- `node -e` parse of the inline `<script>` (no syntax errors).
- Open the live site after deploy and confirm it loads; if data flow changed, test one save and one admin read, then tell the owner to delete any test rows.
