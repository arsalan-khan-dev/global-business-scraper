<div align="center">

<h1>Global Business Scraper</h1>
<h3>Lead Generator — Google Powered v2.0</h3>
<p>Scrape local business data from Google, Google Maps, DuckDuckGo, and Bing — export to CSV, Excel, and Word.</p>

<br/>

[![GitHub Repo](https://img.shields.io/badge/SOURCE-GitHub%20Repo-161b22?style=for-the-badge&logo=github&logoColor=white)](https://github.com/arsalan-khan-dev/global-business-scraper)
[![License](https://img.shields.io/badge/LICENSE-MIT-1565C0?style=for-the-badge)](./LICENSE)
[![Legal Use Only](https://img.shields.io/badge/USE-Ethical%20%26%20Responsible%20Only-1565C0?style=for-the-badge&logo=shield&logoColor=white)](#legal-disclaimer)

<br/>

![Python](https://img.shields.io/badge/Python-3.8%2B-3776AB?style=flat-square&logo=python&logoColor=white)
![Tkinter](https://img.shields.io/badge/GUI-Tkinter%20Desktop%20App-1565C0?style=flat-square)
![Requests](https://img.shields.io/badge/HTTP-Requests%20%2B%20Retry-orange?style=flat-square)
![BeautifulSoup](https://img.shields.io/badge/Parsing-BeautifulSoup4%20%2B%20lxml-2e7d32?style=flat-square)
![openpyxl](https://img.shields.io/badge/Export-Excel%20%2F%20CSV%20%2F%20Word-217346?style=flat-square)
![Threading](https://img.shields.io/badge/Threading-Background%20Scraping-8e44ad?style=flat-square)
![Countries](https://img.shields.io/badge/Coverage-8%20Countries%20%7C%2035%2B%20Categories-1565C0?style=flat-square)

<br/>

> A full-featured, cross-platform desktop application for scraping and exporting local business leads.
> Cascading location dropdowns, 5-source search pipeline, live results table, activity log,
> contact enrichment, and export to CSV, Excel, and Word — all in a single Python file.

</div>

<br/>

---

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [Scraping Pipeline](#scraping-pipeline)
- [Project Structure](#project-structure)
- [Location Coverage](#location-coverage)
- [Business Categories](#business-categories)
- [Data Fields](#data-fields)
- [Export Formats](#export-formats)
- [GUI Overview](#gui-overview)
- [Architecture](#architecture)
- [Dependencies](#dependencies)
- [Getting Started](#getting-started)
- [Legal Disclaimer](#legal-disclaimer)
- [Author](#author)

---

## Overview

**Global Business Scraper** is a production-grade desktop lead generation tool built with Python, Tkinter, and BeautifulSoup4. It queries five search sources in priority order — Google Local, Google Maps, Google Organic, DuckDuckGo, and Bing — to collect real business listings for any category in any supported city. Results are deduplicated, enriched by visiting individual business websites, and exported to structured CSV, Excel, or Word reports.

The entire tool runs as a single Python script with a polished Tkinter GUI. The scraper runs in a background daemon thread so the interface stays fully responsive during long scans.

---

## Features

**5-Source Search Pipeline**
Queries sources in priority order: Google Local (`tbm=lcl`), Google Maps HTML, Google Organic search, DuckDuckGo HTML, and Bing. Each source uses a distinct scraping strategy with graceful fallback if the previous source returns no results.

**Cascading Location Dropdowns**
Country to Province/State to City to Area/Locality. All dropdowns cascade from a built-in `LOCATION_DATA` dictionary covering 8 countries, 25+ provinces, 40+ cities, and 200+ named areas. The Area field is also editable for custom input.

**35+ Business Categories**
Pre-populated dropdown with 35 common business types — from Hospitals and Schools to Bike Shops and Travel Agencies. A custom category field lets users type any search term.

**Contact Information Extraction**
Extracts phone numbers using 6 regex patterns covering international and local formats including Pakistani `+92` and `0xxx-xxxxxxx`, email addresses with spam-domain filtering, and website URLs from both search snippets and business pages.

**Business Page Enrichment**
For the top 12 results, visits each business website directly to extract phone, email, and address — overwriting snippet-level data with richer contact details. Skips social media, aggregator, and map domains automatically.

**Website Status Check**
Performs a live `HEAD` request (with `GET` fallback) for every result's website and records `Yes` or `No` in the `website_status` field. Status checks run with a 6–8 second timeout per URL.

**TTL-Aware Session Management**
Each request uses a `requests.Session` with `HTTPAdapter` retry logic (3 retries, 1.5x backoff, retry on 429/500/502/503/504). User-Agent is rotated from a pool of 6 real browser strings. Random delays (1.5–6 seconds) between requests mimic organic browsing.

**Real-Time Results Table**
`ttk.Treeview` with 9 columns, alternating row colors, column-click sorting, inline filter box, and right-click context menu. Rows with phone or email are highlighted in green. Rows sourced from Google are highlighted in blue.

**Dark Activity Log**
VS Code-styled dark terminal log with colour-coded levels: info (blue), success (teal), warning (yellow), error (red), section (bright blue bold). Timestamped with `HH:MM:SS` prefix. Scrollable and clearable.

**Live Progress Bar**
Green progress bar with percentage label tracks the scraper through 8 named phases: Google Local, Google Maps, Google Organic, DuckDuckGo, Bing, Parse/Dedup, Enrich, Website Check.

**Three Export Formats**
One-click export to CSV, styled Excel workbook with title row and metadata header, or Word document with a formatted table — all saved to a user-chosen path via file dialog.

---

## Scraping Pipeline

```
User clicks "Start Scraping"
          |
          v
  Phase 1 — Google Local (tbm=lcl)          [progress: 5%]
  scrape_google_maps_places()
  Sends 2 query variants if first returns nothing
          |
          v
  Phase 2 — Google Maps HTML                 [progress: 15%]
  scrape_google_maps()
  Strategy A: HTML listing containers
  Strategy B: <script> JS blob lat/lng extraction
          |
          v
  Phase 3 — Google Organic                   [progress: 28%]
  scrape_google()
  Up to 3 query variants
  Parses: div.g / div.MjjYud / div.tF2Cxc
  Also parses Google Local Pack cards
          |
          v
  Phase 4 — DuckDuckGo Fallback              [progress: 45%]
  scrape_duckduckgo()
  html.duckduckgo.com — no JS required
          |
          v
  Phase 5 — Bing Fallback                    [progress: 52%]
  scrape_bing()
  li.b_algo result items
          |
          v
  Phase 6 — Parse and Deduplicate            [progress: 58%]
  sort by source priority (Google Local first)
  normalise name: strip non-word chars + lowercase + set dedup
  skip titles: "wikipedia", "how to", "list of", "top 10", "reviews"
          |
          v
  Phase 7 — Page Enrichment                  [progress: 65-87%]
  scrape_business_page() on top 12 URLs
  extracts phone / email / address from live website
          |
          v
  Phase 8 — Website Status Check             [progress: 90%]
  HEAD request per business website
  Records "Yes" / "No" in website_status field
          |
          v
  Done — results displayed + ready to export [progress: 100%]
```

---

## Project Structure

```
business-scraper/
│
└── business_scraper.py        # Entire application — single file
                               #
                               # Configuration
                               #   LOCATION_DATA          8 countries, 40+ cities
                               #   BUSINESS_CATEGORIES    35 preset categories
                               #   USER_AGENTS            6 real browser strings
                               #   GOOGLE_HEADERS_POOL    Full header sets per UA
                               #
                               # Utilities
                               #   get_session()          Retry-enabled session
                               #   build_query()          Location-aware query builder
                               #   extract_phone()        6 regex phone patterns
                               #   extract_email()        Email + spam filter
                               #   extract_website()      URL from soup/text
                               #   check_website_status() HEAD/GET live check
                               #
                               # Scrapers
                               #   scrape_google()                Organic results
                               #   scrape_google_maps()           Maps HTML + JS
                               #   scrape_google_maps_places()    Local tab
                               #   scrape_duckduckgo()            HTML fallback
                               #   scrape_bing()                  HTML fallback
                               #   scrape_business_page()         Direct page visit
                               #
                               # Processing
                               #   parse_search_result_to_business()
                               #   run_scraper()          8-phase orchestrator
                               #
                               # Exports
                               #   export_csv()           UTF-8 BOM CSV
                               #   export_excel()         Styled openpyxl workbook
                               #   export_word()          python-docx report
                               #
                               # GUI
                               #   BusinessScraperApp     Main Tkinter window
                               #   _build_left_panel()    Filters + controls
                               #   _build_right_panel()   Results + log tabs
                               #   _build_results_tab()   Treeview table
                               #   _build_log_tab()       Dark terminal log
```

---

## Location Coverage

| Country | Provinces / States | Sample Cities |
| --- | --- | --- |
| Pakistan | KPK, Punjab, Sindh, Balochistan, Islamabad | Peshawar, Lahore, Karachi, Quetta, Rawalpindi, Islamabad |
| United States | California, New York, Texas | Los Angeles, San Francisco, New York City, Houston, Austin |
| United Kingdom | England, Scotland, Wales | London, Manchester, Birmingham, Leeds, Edinburgh, Cardiff |
| India | Maharashtra, Delhi, Karnataka, Tamil Nadu | Mumbai, Pune, New Delhi, Bangalore, Chennai |
| UAE | Dubai, Abu Dhabi, Sharjah | Downtown Dubai, Deira, Khalidiyah, Corniche |
| Canada | Ontario, British Columbia, Quebec | Toronto, Ottawa, Vancouver, Montreal |
| Australia | NSW, Victoria, Queensland | Sydney, Melbourne, Brisbane |
| Germany | Bavaria, Berlin, Hamburg | Munich, Nuremberg, Berlin, Hamburg |
| Saudi Arabia | Riyadh, Makkah, Eastern Province | Riyadh, Jeddah, Dammam, Al Khobar |

Each city includes 4–8 named area/locality options for hyper-local targeting.

---

## Business Categories

35 preset categories are included:

```
Bike Shops          Car Showrooms       Schools             Colleges
Universities        Pharmacies          Hospitals           Restaurants
Charity Orgs        Electronics Shops   Clothing Stores     Bakeries
Grocery Stores      Hotels              Gyms & Fitness      Beauty Salons
Barber Shops        Dentists            Law Firms           Accounting Firms
Real Estate         Banks               Mobile Phone Shops  Computer Repair
Petrol Stations     Auto Repair         Furniture Stores    Jewelry Stores
Bookstores          Cafes               Printing Shops      Tailors
Opticians           Veterinary Clinics  Travel Agencies
```

Any custom search term can be entered in the custom category field to search outside the preset list.

---

## Data Fields

Each scraped business record contains 9 fields:

| Field | Description |
| --- | --- |
| `name` | Business name (max 80 characters) |
| `category` | Search category used to find this business |
| `address` | Street address extracted from snippet, business page, or location string |
| `phone` | Phone number matched by 6 regex patterns (international and local formats) |
| `email` | Email address extracted from snippet or business page, spam domains excluded |
| `website` | Business website URL — HTTP/HTTPS only, social media and maps domains excluded |
| `website_status` | `Yes` if website returns HTTP status below 400, otherwise `No` |
| `maps_link` | Google Maps search URL constructed from business name + location |
| `source` | Which scraper found this: Google Local / Google Maps / Google / DuckDuckGo / Bing |

---

## Export Formats

**CSV** — UTF-8 BOM encoded for Excel compatibility. All 9 fields as column headers. Opens directly in Excel, Google Sheets, or any spreadsheet tool.

**Excel (.xlsx)** — Built with `openpyxl`. Includes a merged title row with deep blue background (`#1a237e`), a 4-row metadata block (generated timestamp, location, category, total count), styled header row with white text on navy blue (`#283593`), alternating row fills (white / light indigo `#E8EAF6`), column widths tuned per field, and frozen panes starting at the first data row.

**Word (.docx)** — Built with `python-docx`. Includes a centred heading with indigo colour, a 4-row metadata table, a full business listings table with bold headers, and a greyed footer with the generation date. Page margins set to 2.5 cm on each side.

---

## GUI Overview

**Left Panel — Search Controls**
Scrollable panel with cascading dropdowns for Country → Province → City → Area, category selector with 35 presets, custom category text input, max results selector (25 / 50 / 100 / 200 / 500), Start Scraping and Stop buttons, three export buttons (CSV / Excel / Word), and a Clear Results button.

**Right Panel — Results and Log**
Two-tab layout. The Results tab contains a live filter input and a 9-column `Treeview` table with column-click sorting. Right-clicking any row opens a context menu: Copy Name, Copy Phone, Copy Email, Open Website, Open Maps. Double-clicking a row opens the business website in the default browser.

The Activity Log tab is a dark `ScrolledText` widget styled like a terminal, with colour-coded output by log level and a Clear Log button.

A green progress bar with percentage label spans the bottom of the right panel.

**Colour System**

```python
C_SIDEBAR    = "#1a237e"   # Deep indigo — banner and section headers
C_BTN        = "#3949ab"   # Button fill
C_SUCCESS    = "#2e7d32"   # Progress bar and success status
C_ROW_ALT    = "#E8EAF6"   # Alternating table row fill (light indigo)
C_HEADER_BG  = "#283593"   # Treeview column header background
```

---

## Architecture

`business_scraper.py` is organized into 5 logical layers:

```
Layer 1 — Configuration
  LOCATION_DATA          Nested dict: country → province → city → [areas]
  BUSINESS_CATEGORIES    Flat list of 35 preset search terms
  USER_AGENTS            6 real browser UA strings for rotation
  GOOGLE_HEADERS_POOL    Full HTTP header dicts generated per UA

Layer 2 — HTTP Utilities
  get_session()          requests.Session + HTTPAdapter(Retry(3, backoff=1.5))
  build_query()          Constructs location-aware search string
  extract_phone()        Regex chain: +92, 0xxx, international, 11-digit
  extract_email()        Regex + skip list (noreply, test, sentry, domain, example)
  extract_website()      BeautifulSoup <a href> scan + regex fallback
  check_website_status() HEAD request with GET fallback, 6-8s timeout

Layer 3 — Scrapers (5 sources + enrichment)
  scrape_google_maps_places()    Google Local tab (tbm=lcl), 8 CSS selectors tried
  scrape_google_maps()           Maps HTML listing containers + JS blob extraction
  scrape_google()                Organic div.g results + Local Pack cards
  scrape_duckduckgo()            html.duckduckgo.com — .result__body parsing
  scrape_bing()                  bing.com — li.b_algo item parsing
  scrape_business_page()         Direct GET on business URL, extracts 4 fields

Layer 4 — Orchestration and Processing
  run_scraper()                  8-phase pipeline with 4 callback hooks
  parse_search_result_to_business()  Normalises raw dict → 9-field business record
  Dedup: re.sub(r'\W+', '', name.lower()) → set membership
  Priority sort before dedup: Google Local=0, Maps=1, Maps JS=2, Google=3, DDG=4, Bing=5

Layer 5 — GUI (BusinessScraperApp extends tk.Tk)
  Thread model: daemon Thread(target=run_scraper) for non-blocking UI
  Thread-to-GUI bridge: self.after(0, lambda: ...) for all updates
  Exports: filedialog.asksaveasfilename() → export_csv / export_excel / export_word
```

---

## Dependencies

| Library | Purpose | Required |
| --- | --- | --- |
| `requests` | HTTP requests, session management, retry adapter | Yes — core |
| `beautifulsoup4` | HTML parsing for all 5 scraper sources | Yes — core |
| `lxml` | Fast HTML parser backend for BeautifulSoup4 | Yes — core |
| `openpyxl` | Styled Excel `.xlsx` export | Optional |
| `python-docx` | Word `.docx` export | Optional |
| `pandas` | Detected at startup, reserved for future CSV processing | Optional |
| `tkinter` | Desktop GUI framework | Included with Python |

Install all dependencies at once:

```bash
pip install requests beautifulsoup4 lxml openpyxl python-docx
```

The app starts and scrapes successfully with missing optional packages — it disables the relevant export button and logs a `pip install` hint in the activity log on startup.

---

## Getting Started

**Requirements:** Python 3.8 or higher. Tkinter is included with all standard Python distributions on Windows and macOS. On Linux install `python3-tk` via your package manager.

```bash
# 1. Clone the repository
https://github.com/arsalan-khan-dev/global-business-scraper.git
cd business-scraper

# 2. Install dependencies
pip install requests beautifulsoup4 lxml openpyxl python-docx

# 3. Run the application
python business_scraper.py
```

**Step-by-step usage:**

```
1. Select Country, then Province, then City from the cascading dropdowns
2. Optionally select a specific Area/Locality for hyper-local results
3. Select a Business Category from the 35 presets, or type a custom one
4. Set Max Results — 50 is a good starting point for most searches
5. Click "Start Scraping" and watch results appear live in the table
6. Use the filter box to search within results by any field
7. Right-click any row to copy contact details or open the website
8. Click Export CSV / Export Excel / Export Word to save your lead list
```

**Tips for better results:**
- Start broad (city only) before narrowing to a specific area
- Google may rate-limit aggressive scans — the built-in delays help but are not a guarantee
- The enrichment phase visits up to 12 websites and adds 30–90 seconds to total runtime
- For faster scans set max results to 25 and use the `--no-enrich` mindset (stop early)
- If Google results are sparse, DuckDuckGo and Bing fallbacks usually fill the gap

---

## Legal Disclaimer

> This tool is intended for ethical, responsible use only.
>
> Web scraping public search engines may be subject to their Terms of Service. Only use this tool for lawful purposes such as market research, lead generation for your own business, or academic study. Do not use it to harvest data for spam, unsolicited outreach, or any activity that violates applicable law including GDPR, Pakistan PECA 2016, or CAN-SPAM Act.
>
> The author accepts no liability for misuse of this tool.

---

<div align="center">

**Arsalan Khan**

[![GitHub](https://img.shields.io/badge/GitHub-arsalan--khan--dev-161b22?style=for-the-badge&logo=github&logoColor=white)](https://github.com/arsalan-khan-dev)
[![Repository](https://img.shields.io/badge/Repo-business--scraper-1565C0?style=for-the-badge&logo=github&logoColor=white)](https://github.com/arsalan-khan-dev/global-business-scraper)

<br/>

*Built with precision. Engineered for lead generation.*

---

<sub>© 2025 Global Business Scraper · Built by Arsalan Khan · For ethical use only</sub>

</div>
