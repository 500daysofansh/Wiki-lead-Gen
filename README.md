# Lead Generation Automation — Python Internship Assignment

## Approach & Tools

**Approach:**  
The script runs a two-source collection pipeline — a curated seed dataset of 30 prominent India-based NGOs/nonprofits, supplemented by a live Wikipedia table scrape. Data is cleaned (deduplication, null-filling, string normalisation) and enriched with auto-generated probable contact emails derived from organisation name + domain. The final dataset is exported to a formatted, colour-coded Excel workbook with a Leads sheet and a Summary sheet.

**Tools Used:**
| Tool | Purpose |
|------|---------|
| `requests` + `BeautifulSoup` | Web scraping (Wikipedia table) |
| `pandas` | Data cleaning, deduplication, transformation |
| `openpyxl` | Excel creation with full formatting |
| `schedule` | Bonus: daily scheduled run at 08:00 AM |

---

## Project Structure

```
lead_gen_automation/
├── lead_scraper.py      # Main automation script
├── leads_output.xlsx    # Generated output file
└── README.md            # This file
```

---

## How to Run

### One-time run
```bash
pip install requests beautifulsoup4 openpyxl pandas schedule
python lead_scraper.py
```

### Scheduled daily run (Bonus)
```bash
python lead_scraper.py --schedule
```
This runs the pipeline every day at 08:00 AM automatically.

---

## Features

- ✅ 30 leads collected (NGOs/nonprofits across India)
- ✅ Fields: Name, Email, Website, LinkedIn, Location, Source, Status, Collected On
- ✅ Deduplication by Name
- ✅ Missing value handling
- ✅ **Bonus:** Probable email format generation (`info@domain.com`)
- ✅ **Bonus:** Scheduled daily automation trigger (`--schedule` flag)
- ✅ Formatted Excel output (header colours, alternating rows, auto-filter, freeze panes)
- ✅ Summary sheet with key metrics

---

## Output Preview

The Excel file contains two sheets:
1. **Leads** — Full dataset with all fields, formatted table
2. **Summary** — Quick stats: total leads, email/website/LinkedIn coverage, run date
