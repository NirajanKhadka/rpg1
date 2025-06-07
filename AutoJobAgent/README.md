# AutoJobAgent 🤖📄🚀  

Automate the boring parts of job-hunting:  
*One CLI → scrape new jobs → LLM-tailor your résumé & cover letter → convert to PDF → auto-submit to Workday / iCIMS / Greenhouse → live dashboard → pause / resume safely – zero duplicates or missing keywords.*

---

## 1. What is AutoJobAgent?

AutoJobAgent is a **Windows-first** Python toolkit that:

1. Searches Indeed / Eluta (and more soon) for fresh postings that match your keywords.  
2. Calls a local **Ollama** LLM (`mistral:7b` by default) to weave missing keywords into your résumé & cover letter.  
3. Converts DOCX → PDF via Word COM (100 % fidelity).  
4. Detects the target ATS (Workday, iCIMS, Greenhouse; Lever stub) and drives a **Playwright** Chromium session to upload forms, answer screening questions and click *Submit*.  
5. Stores every attempt in `output/logs/applications.xlsx` and streams stats to a FastAPI/WebSocket dashboard.  
6. Can be paused/resumed at any time without losing progress—no duplicate applications.

---

## 2. Requirements

| Layer | Version / Detail | Why |
|-------|------------------|-----|
| **OS** | Windows 10/11 64-bit | Word automation for fast, reliable PDF export |
| **Microsoft Word** | 2016 or newer (bit-ness must match Python) | FileFormat 17 conversion |
| **Python** | 3.10 ×64 | Compatible with Playwright wheels |
| **Browser** | `playwright install chromium` (headed) | Lets you jump in if CAPTCHAs appear |
| **LLM** | Ollama daemon with `mistral:7b` pulled | No cloud keys required |
| **Python packages** | see `requirements.txt` | Scraping, PDF, API, UI |

> 🐍 *Optional*: `conda create -n auto_job python=3.10`

---

## 3. Installation

```bash
git clone https://github.com/your-org/AutoJobAgent.git
cd AutoJobAgent

# Create & activate env (venv or conda)
python -m venv .venv && .\.venv\Scripts\activate

pip install -r requirements.txt
playwright install chromium          # one-time
ollama pull mistral:7b               # one-time
```

Word COM is detected automatically; if missing, PDFs are skipped and DOCX files are uploaded instead.

---

## 4. Setup

### 4.1 Profiles

Copy `profiles/Nirajan` as a template for each person:

```text
profiles/
└── Alice/
    ├── Alice.json
    ├── Alice_Resume.docx
    └── Alice_CoverLetter.docx
```

Minimal **JSON schema** (keywords are mandatory):

```jsonc
{
  "name": "Alice Smith",
  "location": "Toronto, ON",
  "email": "alice@example.com",
  "phone": "416-555-9999",
  "keywords": ["Data Analyst", "Python", "SQL"],
  "skills": [...],                // optional
  "batch_default": 15,            // optional
  "ollama_model": "mistral:7b"    // optional
}
```

### 4.2 Environment variables (optional)

| Variable | Purpose | Default |
|----------|---------|---------|
| `AUTOJOB_OLLAMA_URL` | custom Ollama endpoint | `http://localhost:11434` |
| `AUTOJOB_HEADLESS`  | `1` to force headless browser | interactive |
| `AUTOJOB_ALLOW_SENIOR` | `1` to apply to senior roles | skip |

---

## 5. Usage

### 5.1 Quick start

```bash
python main.py Nirajan
```

### 5.2 Common flags

| Flag | Description |
|------|-------------|
| `--batch 25` | override batch size |
| `--ats auto|workday|icims|greenhouse|manual` | pick ATS or prepare docs only |
| `--verbose` | extra console logs |
| `--allow-senior` | ignore senior-title filter |
| `--headless` | hide the browser |

### 5.3 Dashboard

```bash
uvicorn dashboard_api:app --reload   # dev mode
```

Navigate to <http://localhost:8000>.  
Live stats, latest logs, and **Pause / Resume** buttons are powered by WebSockets.

---

## 6. Directory layout

```
AutoJobAgent/
│
├── main.py                 ← Rich CLI entrypoint
├── job_scraper.py          ← multi-site generator
│   └── scrapers/indeed.py …
├── document_generator.py   ← LLM + DOCX + PDF
├── ats/                    ← Playwright submitters
│   ├── workday.py
│   ├── icims.py
│   └── greenhouse.py
├── dashboard_api.py        ← FastAPI + WebSocket
│   └── templates/index.html
├── utils.py                ← PDF, hashing, Excel, helpers
├── profiles/…              ← per-user folders
├── output/                 ← logs + generated PDFs
└── tests/                  ← pytest stubs
```

---

## 7. Configuration & Sessions

A `session.json` is auto-created inside each profile:

```json
{
  "ats": "workday",
  "next_index": 30,
  "done": ["6a41a8f1c8", "2e9cf12bd5"]
}
```

Delete the file to start fresh or keep it to resume exactly where you left off.

---

## 8. Troubleshooting

| Symptom | Cause | Fix |
|---------|-------|-----|
| **`comtypes` or `pywin32` import error** | Wrong Python bit-ness | Install 64-bit Python, match Word |
| Word opens but no PDF | Word not licensed / activation screen | Activate Word or run with `--ats manual` |
| CAPTCHA appears | Site rate-limited | Agent pauses and asks you to solve, then hit *Enter* |
| Excel log locked | You opened the file in Excel | Close Excel or wait – agent retries 5× |
| Duplicate posting | Same hash already logged | Agent auto-skips |
| Senior roles skipped | Title contains “Senior/Lead/Director” | Run with `--allow-senior` |
| Browser too fast | Animations not loaded | Add `--headless` off (default) and slow network throttling if needed |

---

## 9. Planned Features

- Google Jobs & LinkedIn scrapers  
- Lever ATS automation  
- GitHub Actions CI (`pytest`, `flake8`)  
- Docker **dev-container** (without Office)  
- Slack / Discord notifications  

---

## 10. Contributing & License

PRs and issue reports are welcome!  
This project is licensed under the MIT License. See `LICENSE` for details.

Happy auto-applying! 💼✨
