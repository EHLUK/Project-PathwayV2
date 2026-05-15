# P6 XER Project Manager Planning Tool

A Streamlit web app that lets Project Managers, Project Leads and operational teams
interrogate Primavera P6 XER schedules without needing to open P6.

---

## Folder Structure

```
p6_planner/
├── app.py              ← Main application (all pages in one file)
├── requirements.txt    ← Python dependencies
├── README.md           ← This file
└── .streamlit/
    └── config.toml     ← Optional Streamlit theme config
```

---

## Features

| Page | What it does |
|---|---|
| 📊 Project Summary | KPIs, charts, float distribution, WBS breakdown |
| 🔍 Activity Search | Filter and search activities; view detail + logic |
| 🔗 Logic Trace | Trace predecessors / successors with depth levels |
| 🚨 Critical Path Analysis | Full critical path, near-critical, negative float |
| 🎯 Critical Path to Activity | Identify what is driving any milestone or activity |
| 👷 Labour Histogram | Weekly/monthly resource histograms by trade/WBS |
| 🩺 Schedule Health Check | 11 automated quality checks with export |
| 📝 Planning Notes | Upload notes, link to activities, keyword search |
| 📅 Programme Comparison | Compare two XER revisions; date/float movement |
| 📥 Export Reports | Download Excel reports for all data sets |

---

## Running Locally

### 1. Prerequisites
- Python 3.10 or later
- pip

### 2. Create a virtual environment (recommended)
```bash
python -m venv venv
# Windows:
venv\Scripts\activate
# macOS / Linux:
source venv/bin/activate
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

### 4. Run the app
```bash
streamlit run app.py
```

The app will open in your browser at `http://localhost:8501`.

### 5. Export an XER file from P6
In Primavera P6:
- **File → Export → Primavera P6 – Project Management (XER)**
- Select your project and export
- Upload the `.xer` file using the sidebar

---

## Deploying to Streamlit Community Cloud

### 1. Push to GitHub
Create a public (or private) GitHub repository and push:
```bash
git init
git add app.py requirements.txt README.md
git commit -m "Initial commit - P6 Planner Tool"
git branch -M main
git remote add origin https://github.com/YOUR_USERNAME/p6-planner.git
git push -u origin main
```

### 2. Deploy on Streamlit Community Cloud
1. Go to [share.streamlit.io](https://share.streamlit.io)
2. Sign in with your GitHub account
3. Click **New app**
4. Select your repository, branch (`main`), and main file (`app.py`)
5. Click **Deploy**

Streamlit Community Cloud will install dependencies from `requirements.txt` automatically.

### 3. Optional: Add theme config
Create `.streamlit/config.toml`:
```toml
[theme]
primaryColor = "#2563eb"
backgroundColor = "#ffffff"
secondaryBackgroundColor = "#f0f4f8"
textColor = "#1e3a5f"
font = "sans serif"
```

---

## Notes for Users

- **XER file size**: Large programmes (10,000+ activities) may take 5–15 seconds to parse
- **Resource histograms**: Only available if resources were assigned in P6 before export
- **Labour CSV upload**: If no resources in XER, you can upload a separate CSV with columns:
  `task_code, rsrc_name, target_qty, target_start, target_finish`
- **Near-critical threshold**: Adjustable in the sidebar (default 10 working days float)
- **Programme comparison**: Upload two separate XER files on the Comparison page (no main file needed)

---

## Troubleshooting

| Problem | Solution |
|---|---|
| "Cannot decode XER file" | Try re-exporting from P6; check the export used cp1252 or UTF-8 encoding |
| "xerparser failed, using fallback" | Normal — the fallback parser handles most XER files |
| No resource data | Resources were not assigned/exported; use the CSV upload option |
| Dates show as None | Check that the programme was fully scheduled before export |
| App is slow | Large files (5000+ activities) take longer; consider filtering in P6 before export |

---

## Dependencies

| Package | Purpose |
|---|---|
| streamlit | Web app framework |
| pandas | Data manipulation |
| plotly | Interactive charts and Gantt diagrams |
| networkx | Logic network graph traversal |
| openpyxl | Excel export with formatting |
| xlsxwriter | Alternative Excel writer |
| python-docx | Read DOCX planning notes |
| xerparser | Primary P6 XER file parser |
