# Jobright Multi-Category Job Scraper & Salary Dashboard

A self-hosted pipeline that scrapes live job postings from [Jobright.ai](https://jobright.ai) for multiple job categories — **Data Scientist, Data Analyst, Machine Learning Engineer, and Statistician** by default — normalizes each posting's required major, job role, and seniority level into a consistent taxonomy, and presents the combined dataset through an interactive Streamlit dashboard.

Job categories are **not hardcoded** — they live in an editable JSON config file, so you can add or remove categories at any time without touching code.

---

## Features

- 🔎 **Automated scraping** of Jobright job listings via Playwright, with resumable, autosaving progress (safe to stop and restart)
- 🏷️ **Canonical taxonomy mapping** — normalizes messy raw text (e.g. "Info Systems" / "MIS" / "Information Systems") into consistent major, role, and seniority buckets using alias rules, wildcard patterns, and fuzzy matching
- 🗂️ **Dynamic category system** — add, remove, or edit job categories via a JSON config, the CLI, or the GUI, with no code changes required
- 🔁 **Multi-category merge** — combines all scraped categories into one Excel file, tagged by category, de-duplicated by job URL
- 🖥️ **Desktop GUI** (Tkinter) — scrape, merge, and launch the dashboard from one window
- 📊 **Interactive dashboard** (Streamlit + Plotly) — three tabs covering raw data exploration, branch/role salary summaries, and overall statistics
- 🔐 **One-time login** — your Jobright session is saved and reused across categories, so you only log in once per machine

---

## Project Structure

```
project-root/
├── code/
│   ├── pipelinescrapper_core.py       # Core Playwright scraping engine
│   ├── canonical_postprocess.py       # Major / role / seniority taxonomy mapping
│   ├── pipeline_categories.py         # Dynamic category config + orchestration
│   ├── scrape_category.py             # CLI: scrape / add / remove / list categories
│   ├── merge_categories.py            # CLI: merge all categories into one Excel
│   ├── gui_runner.py                  # Tkinter desktop GUI
│   ├── dashboard_app.py               # Streamlit dashboard
│   └── setup_and_run.py               # Installer + single entry point for all commands
├── input/
│   ├── canonical_majors_custom_full.xlsx   # Reference list of canonical majors
│   └── categories_config.json              # Auto-created — editable job category list
└── output/
    ├── <category_key>/
    │   ├── jobright_jobs.xlsx          # Per-category scraped + canonicalized data
    │   ├── jobs_raw.jsonl               # Raw scrape log (for resuming)
    │   └── state.json                   # Scrape progress state (for resuming)
    ├── dom_snapshots/
    │   └── jobright_login_session.json  # Saved login session (DO NOT COMMIT — see below)
    ├── jobright_jobs.xlsx               # Combined dataset across all categories
    └── branch_salary_dashboard_ready.xlsx  # Generated summary stats
```

---

## Setup

### 1. Requirements
- Python 3.10+
- A Jobright.ai account (free tier works)

### 2. Install dependencies
```bash
python code/setup_and_run.py install
```
Installs `pandas`, `openpyxl`, `playwright`, `streamlit`, `plotly`, `beautifulsoup4`, and downloads the Playwright Chromium browser binary.

---

## Usage

### Command line

**List available categories:**
```bash
python code/setup_and_run.py scrape list
```

**Scrape a category (target job count is optional, default 150):**
```bash
python code/setup_and_run.py scrape data_scientist 150
python code/setup_and_run.py scrape data_analyst 150
python code/setup_and_run.py scrape machine_learning_engineer 150
python code/setup_and_run.py scrape statistician 150
```
A Chromium window opens — log into Jobright the first time (your session is reused automatically after that). Come back to the terminal and press Enter to continue scraping.

**Add a new category without editing code:**
```bash
python code/scrape_category.py add data_engineer "Data Engineer" "https://jobright.ai/jobs/data-engineer-jobs-in-united-states" "data engineer"
```

**Remove a category:**
```bash
python code/scrape_category.py remove data_engineer
```

**Merge all scraped categories into one file:**
```bash
python code/setup_and_run.py merge
```

**Launch the dashboard:**
```bash
python code/setup_and_run.py dash
```
Opens at `http://localhost:8501`.

### Desktop GUI

```bash
python code/setup_and_run.py gui
```
Pick a category from the dropdown (or use **+ Add** / **− Remove** to manage the list), choose a job count, click **Start Scraping**, log in, click **Continue after login**, then **Merge All Categories** and **Open Dashboard** when done.

---

## Dashboard Tabs

| Tab | Contents |
|---|---|
| **Main Dashboard** | Raw job listings with filters (category, company, work model, role type, seniority), branch-metric bar charts, and a configurable multi-axis chart builder |
| **Branch Salary Summary** | Count, average/median salary, salary range, and std. deviation per canonical major, role, or seniority bucket — bar, pie, scatter, and comparison views |
| **Overall Statistics** | KPI cards, jobs-by-category breakdown, per-category salary comparison, salary distribution histogram, work-model/seniority splits, top hiring companies |

---

## Configuring Job Categories

Categories live in `input/categories_config.json`:

```json
{
  "data_scientist": {
    "label": "Data Scientist",
    "url": "https://jobright.ai/jobs/data-scientist-jobs-in-united-states",
    "role_selected": "data scientist"
  }
}
```

Edit this file directly, or use the CLI/GUI methods above — changes take effect on the next run without restarting anything.

---

## Notes & Limitations

- Jobright requires an authenticated session; at least one manual login is needed before scraping can proceed.
- Scraping is scoped to United States postings by default (change the category URL to target other regions).
- Canonical major mapping quality depends on the completeness of `input/canonical_majors_custom_full.xlsx`.
- The scrape loop stops automatically after ~1500 scrolls with no new postings.

---

## ⚠️ Before pushing to a public repo

Do **not** commit:
- `output/dom_snapshots/jobright_login_session.json` — this is your saved login session/cookies
- `output/**/*.xlsx`, `*.jsonl`, `*.json` scrape data — regenerable and can be large
- `__pycache__/`

A suggested `.gitignore`:
```
__pycache__/
*.pyc
output/
.venv/
venv/
```

---

## License

Add your preferred license here (e.g. MIT).
