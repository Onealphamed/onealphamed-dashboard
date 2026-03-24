# OnealphaMed — Project Profit Dashboard

Interactive profit dashboard that auto-updates whenever the Excel data file changes.

## How it works

```
Project Profit.xlsx  →  generate_dashboard.py  →  OnealphaMed_Dashboard.html
```

1. You edit **`Project Profit.xlsx`** with new data
2. Push it to GitHub
3. GitHub automatically runs the Python script
4. The HTML dashboard is updated and committed back
5. GitHub Pages serves the live dashboard URL

---

## Repository structure

```
/
├── Project Profit.xlsx          ← Your data file (edit this)
├── generate_dashboard.py        ← The generator script (do not edit)
├── requirements.txt             ← Python dependencies
├── OnealphaMed_Dashboard.html   ← Auto-generated dashboard (do not edit manually)
└── .github/
    └── workflows/
        └── update_dashboard.yml ← GitHub Actions automation
```

---

## Setup (one-time)

### Step 1 — Create a GitHub repository
1. Go to [github.com](https://github.com) → **New repository**
2. Name it e.g. `onealphamed-dashboard`
3. Set it to **Public** (required for free GitHub Pages)

### Step 2 — Upload all files
Upload these files to your repo:
- `Project Profit.xlsx`
- `generate_dashboard.py`
- `requirements.txt`
- `OnealphaMed_Dashboard.html`
- `.github/workflows/update_dashboard.yml`

### Step 3 — Enable GitHub Pages
1. Go to your repo → **Settings** → **Pages**
2. Under *Source*, select **Deploy from a branch**
3. Choose **main** branch, **/ (root)** folder
4. Click **Save**

Your dashboard will be live at:
`https://YOUR-USERNAME.github.io/onealphamed-dashboard/OnealphaMed_Dashboard.html`

---

## Updating the dashboard

1. Open `Project Profit.xlsx` and make your changes
2. Save the file
3. Go to GitHub → drag & drop the updated Excel file → **Commit changes**
4. GitHub Actions runs automatically (takes ~30 seconds)
5. The dashboard refreshes — done!

> **Tip:** You can also trigger a manual update anytime from
> GitHub repo → **Actions** tab → **Update Dashboard** → **Run workflow**

---

## What the dashboard shows

| Tab | Contents |
|-----|----------|
| 📊 Overview | Grand total KPIs, monthly bar chart, company pie chart, filterable margin trend |
| Per company (×12) | KPI cards, monthly bar chart, project pie chart, month-wise table, project detail table |
| 🔧 Vendor Analysis | Medical Writing / Webinar-Tech / Events breakdown by month and company |

**Companies tracked:** Hetero · Bayer · Lupin · P&G · Cipla · Aurobindo · NovoNordisk · Zydus · KOITA · Amneal · Resmed · Sun Pharma
