# Google Maps Business Scraper

A powerful scraper that extracts business leads from Google Maps — including **emails**, **social media links**, **phone numbers**, and more.

## 🚀 Features

- **Email Extraction** — Finds business emails from their website (homepage + /contact pages), Google Maps listing, or via Google search
- **Social Media Links** — Scrapes Instagram, Facebook, Twitter/X, LinkedIn, TikTok, YouTube from business websites
- **Flexible Filtering** — Scrape businesses WITH or WITHOUT websites
- **No-City Mode** — Scrape at the country or region level (e.g. "restaurants in UK")
- **No-Niche Mode** — Provide only a location to discover diverse business types automatically
- **Cross-Run Deduplication** — Never scrapes the same business twice (persisted in `seen_urls.json`)
- **Excel & CSV Output** — Auto-formatted `.xlsx` with auto-column-widths, or `.csv`
- **Crash-Resistant** — Handles locked Excel files, timeouts, and bad websites gracefully

## 📋 Requirements

```bash
pip install playwright pandas openpyxl
python -m playwright install chromium
```

## 📖 Usage & Commands

### Basic Commands

| Command | Description |
|---------|-------------|
| `--query` | Business type / niche (e.g. `"restaurants"`, `"plumbers"`). Optional. |
| `--location` | Country, region, or city (e.g. `"UK"`, `"India"`, `"Chicago IL"`). Optional. |
| `--max` | Maximum number of leads to collect (default: 60) |
| `--output` | Output file path (default: `no_website_businesses.xlsx`) |
| `--with-websites` | Collect businesses that HAVE a website (default: those WITHOUT) |
| `--reset-seen` | Clear the dedup cache and start fresh |

> **Note:** At least one of `--query` or `--location` is required.

---

### 🔍 Scrape by Niche + Location

Find businesses of a specific type in a specific area:

```bash
# Restaurants in India (without websites)
python gmaps_scraper.py --query "restaurants" --location "India"

# Plumbers in Chicago (max 50 results)
python gmaps_scraper.py --query "plumbers" --location "Chicago IL" --max 50

# Dentists in Kosovo (WITH websites — also scrapes email & socials from site)
python gmaps_scraper.py --query "dentists" --location "Kosovo" --with-websites
```

### 🌍 Scrape by Country/Region (No City Needed)

Just provide a country or region name — no specific city required:

```bash
# Restaurants across the UK
python gmaps_scraper.py --query "restaurants" --location "UK" --max 50

# Hotels in Germany
python gmaps_scraper.py --query "hotels" --location "Germany" --max 30
```

### 🎲 Scrape Location Only (No Niche — Diverse Businesses)

Omit `--query` to automatically discover various business types in a location:

```bash
# Diverse businesses in the UK
python gmaps_scraper.py --location "UK" --max 50

# Diverse businesses in Berlin, with websites + email/social scraping
python gmaps_scraper.py --location "Berlin" --with-websites --max 30
```

This searches multiple broad categories (restaurants, shops, services, stores, salons, clinics, etc.) and combines the results.

### 🌐 With vs Without Websites

```bash
# Businesses WITHOUT websites (default) — still tries to find emails via Google search
python gmaps_scraper.py --query "cafes" --location "Paris" --max 20

# Businesses WITH websites — scrapes emails + social links from the actual site
python gmaps_scraper.py --query "cafes" --location "Paris" --with-websites --max 20
```

### 📁 Custom Output

```bash
# Save to a specific Excel file
python gmaps_scraper.py --query "hotels" --location "Tokyo" --output tokyo_leads.xlsx

# Save as CSV
python gmaps_scraper.py --query "hotels" --location "Tokyo" --output tokyo_leads.csv
```

### 🔄 Reset Deduplication Cache

```bash
# Clear the seen-businesses cache and scrape fresh
python gmaps_scraper.py --query "restaurants" --location "London" --reset-seen
```

## 📊 Output Columns

| Column | Description |
|--------|-------------|
| **Name** | Business name |
| **Category** | Business category/type from Google Maps |
| **Phone** | Phone number |
| **Email** | Business email (from website, Maps listing, or Google search) |
| **Address** | Full address |
| **Rating** | Google Maps star rating |
| **Reviews** | Number of reviews |
| **Website** | Business website URL |
| **Instagram** | Instagram profile link |
| **Facebook** | Facebook page link |
| **Twitter** | Twitter/X profile link |
| **LinkedIn** | LinkedIn company/profile link |
| **TikTok** | TikTok profile link |
| **YouTube** | YouTube channel link |
| **Maps URL** | Direct Google Maps link to the listing |

## 🔧 How Email Extraction Works

The scraper uses a **3-tier approach** to find business emails:

1. **From the business website** (if it has one):
   - Scans the homepage for `mailto:` links and email patterns
   - Also checks `/contact`, `/contact-us`, `/about`, `/about-us` pages
   - Extracts social media links from the site too

2. **From the Google Maps listing page**:
   - Scans for `mailto:` links and email patterns visible on the Maps detail page

3. **Via Google Search** (fallback):
   - Searches `"<business name> <location> email contact"` on Google
   - Extracts any email addresses found in the search results

## ⚠️ Notes

- The scraper runs in **headless Chromium** (no visible browser window)
- Results depend on what Google Maps returns — larger locations may show limited results
- Email extraction is best-effort; not all businesses have publicly listed emails
- Social media links are only found when businesses list them on their website
- The dedup cache (`seen_urls.json`) prevents re-scraping the same business across runs — use `--reset-seen` to start fresh

## 🔎 External Email Source

The Google Maps scraper stays unchanged. For a separate source focused on business emails without websites, use the new Yellow Pages scraper:

```bash
python yellowpages_email_scraper.py --query "plumbers" --location "Chicago IL"
python yellowpages_email_scraper.py --location "Chicago IL" --max 50
python yellowpages_email_scraper.py --query "dentists" --location "New York, NY" --include-websites
```

This scraper targets Yellow Pages profile pages, keeps results that expose an email address, and filters out listings that already have a website unless you pass `--include-websites`.
