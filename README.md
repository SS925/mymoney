# MyMoney

Personal finance app for your phone, plus coaching receipts, in one installable web app.

**Open:** https://ss925.github.io/mymoney/ (then *Add to Home screen*)

## What's inside

| Screen | What it shows |
|---|---|
| **Home** | Net worth, money available, left over, loans, savings rate, *within net pay* limit, 12-month chart, health score, alerts |
| **Activity** | Every transaction from Excel: search, type, year, month and category filters; tap for full bank text |
| **＋ Add** | Cash / non-bank entries. They wait in GitHub until Excel imports them into Tracking |
| **Budget** | *Net-pay plan* (editable limits that fit your net pay) or *Excel targets* |
| **More** | **Coaching receipts**, Retirement, phone notifications, and every Excel sheet (Sections / Tiles / Excel grid) |
| **Coaching receipts** (`receipts.html`) | Make a receipt, share/download the PDF, copy as text, past receipts, monthly and all-time totals |

Long-press the app icon (Android) for the **New receipt** and **Add entry** shortcuts.

## Where the data lives (all private)

| Data | Private repo | Written by |
|---|---|---|
| Figures, sheets, phone entries, net-pay plan, notification keys | `SS925/-mymoney-data-` (`data.json`, `sheets.json`, `inbox/`, `plan.json`, `push/`) | Excel's AppLink macro and the app |
| Receipts and your receipt details | `SS925/coaching-receipts-data` (`receipts.json`) | The receipts screen |

This public repo holds only the app code, never your figures or receipts.

**One GitHub key for both:** in GitHub → Settings → Developer settings → Fine-grained tokens, edit your MyMoney token and add
`coaching-receipts-data` under *Repository access* (Contents: Read and write). The receipts screen then connects by itself
using MyMoney's key. You can also paste a separate key in *Coaching receipts → My details*.

Receipts made earlier in the standalone Coaching Receipt Maker on the same phone appear automatically (both apps are on
`ss925.github.io`), and everything in `receipts.json` is downloaded on the first connect.

## Files

| File | Purpose |
|---|---|
| `index.html`, `app.js`, `styles.css` | MyMoney screens |
| `receipts.html` | Coaching receipts (single file; uses html2canvas + jsPDF from cdnjs for the PDF) |
| `sw.js` | Offline support and phone notifications |
| `manifest.webmanifest`, `icon*` | Install on the home screen, app shortcuts |

The Excel macro (`AppLink.bas`) and the notification checker (`notify.yml`, which belongs in the **private** data repo at
`.github/workflows/notify.yml`) are kept with the setup notes, not in this repo.

## Run locally

No build step. Serve the folder over http and open it:

```
python3 -m http.server 8000
# open http://localhost:8000/
```

## Deploy

GitHub Pages: *Settings → Pages → Deploy from a branch → main → / (root)*. After changing files, bump `CACHE` in `sw.js`
so phones pick up the new version, then close and reopen the app.
