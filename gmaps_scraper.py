"""
Google Maps Scraper — Businesses With or Without Websites
=========================================================
Requirements:
    pip install playwright pandas openpyxl
    python -m playwright install chromium

Usage:
    # Businesses WITHOUT a website (default)
    python gmaps_scraper.py --query "restaurants" --location "India"
    python gmaps_scraper.py --query "plumbers" --location "Chicago IL" --max 50

    # Businesses WITH a website (also scrapes email & social links from site)
    python gmaps_scraper.py --query "dentists" --location "Kosovo" --with-websites
    python gmaps_scraper.py --query "hotels" --location "Tokyo" --with-websites --output tokyo_leads.xlsx

    # Country/region-level (no specific city)
    python gmaps_scraper.py --query "restaurants" --location "UK" --max 50

    # Location only — no niche (scrapes diverse business types)
    python gmaps_scraper.py --location "UK" --max 50
    python gmaps_scraper.py --location "Berlin" --with-websites --max 30

Features:
    • --with-websites flag  — collect businesses that DO have a website instead
    • Email & social links  — scrapes email, Instagram, Facebook, Twitter, LinkedIn from website
    • Flexible queries      — location is optional, query is optional
    • No-niche mode         — provide only --location to scrape diverse business types
    • Cross-run dedup       — never scrapes the same business twice (seen_urls.json)
    • Auto-renames output   — if .xlsx is open in Excel (PermissionError fix)
    • Summary printed at the end
"""

import argparse
import asyncio
import json
import os
import re
import urllib.parse
from datetime import datetime
import pandas as pd
from playwright.async_api import async_playwright

# ──────────────────────────────────────────────
# CONFIG
# ──────────────────────────────────────────────
DEFAULT_MAX     = 60
SCROLL_PAUSE    = 1.5
OUTPUT_FILE     = "no_website_businesses.xlsx"
SEEN_FILE       = "seen_urls.json"          # persists seen place IDs across runs
PROGRESS_FILE   = "query_progress.json"     # tracks completed queries for resume

# ──────────────────────────────────────────────
# RICHEST CITIES BY COUNTRY MAPPING
# ──────────────────────────────────────────────
RICHEST_CITIES = {
    "india": [
        "Mumbai", "Delhi", "Bengaluru", "Hyderabad", "Chennai",
        "Kolkata", "Pune", "Ahmedabad", "Surat", "Visakhapatnam"
    ],
    "usa": [
        "New York City", "Los Angeles", "Chicago", "San Francisco", "Houston",
        "Dallas", "Seattle", "Boston", "Washington DC", "Miami"
    ],
    "united states": [
        "New York City", "Los Angeles", "Chicago", "San Francisco", "Houston",
        "Dallas", "Seattle", "Boston", "Washington DC", "Miami"
    ],
    "uk": [
        "London", "Edinburgh", "Manchester", "Birmingham", "Bristol",
        "Glasgow", "Leeds", "Liverpool", "Newcastle", "Sheffield"
    ],
    "united kingdom": [
        "London", "Edinburgh", "Manchester", "Birmingham", "Bristol",
        "Glasgow", "Leeds", "Liverpool", "Newcastle", "Sheffield"
    ],
    "canada": [
        "Toronto", "Montreal", "Vancouver", "Calgary", "Edmonton",
        "Ottawa", "Winnipeg", "Quebec City", "Hamilton", "Kitchener"
    ],
    "australia": [
        "Sydney", "Melbourne", "Brisbane", "Perth", "Adelaide",
        "Gold Coast", "Newcastle", "Canberra", "Sunshine Coast", "Wollongong"
    ],
    "germany": [
        "Munich", "Frankfurt", "Hamburg", "Stuttgart", "Dusseldorf",
        "Berlin", "Cologne", "Hannover", "Nuremberg", "Bremen"
    ],
    "france": [
        "Paris", "Lyon", "Marseille", "Toulouse", "Bordeaux",
        "Lille", "Nice", "Nantes", "Strasbourg", "Rennes"
    ],
    "italy": [
        "Milan", "Rome", "Turin", "Bologna", "Florence",
        "Venice", "Genoa", "Naples", "Verona", "Padua"
    ],
    "spain": [
        "Madrid", "Barcelona", "Valencia", "Seville", "Bilbao",
        "Zaragoza", "Malaga", "Murcia", "Palma", "Las Palmas"
    ]
}

# Broad search terms for no-niche mode (discovers diverse businesses)
BROAD_QUERIES = [
    "restaurants",
    "shops",
    "services",
    "businesses",
    "stores",
    "salons",
    "clinics",
    "gyms",
    "bakeries",
    "pharmacies",
    "mechanics",
    "hotels",
    "cafes",
    "supermarkets",
    "plumbers",
    "real estate",
    "barber",
    "dentist",
    "electrician"
]

# Social media domain patterns for extraction
SOCIAL_PATTERNS = {
    "Instagram": re.compile(r"https?://(?:www\.)?instagram\.com/[^/\s\"'?#]+", re.I),
    "Facebook":  re.compile(r"https?://(?:www\.)?facebook\.com/[^/\s\"'?#]+", re.I),
    "Twitter":   re.compile(r"https?://(?:www\.)?(?:twitter\.com|x\.com)/[^/\s\"'?#]+", re.I),
    "LinkedIn":  re.compile(r"https?://(?:www\.)?linkedin\.com/(?:company|in)/[^/\s\"'?#]+", re.I),
    "TikTok":    re.compile(r"https?://(?:www\.)?tiktok\.com/@[^/\s\"'?#]+", re.I),
    "YouTube":   re.compile(r"https?://(?:www\.)?youtube\.com/(?:@|c/|channel/)[^/\s\"'?#]+", re.I),
}

EMAIL_REGEX = re.compile(
    r"[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}",
    re.I,
)

# Emails to ignore (common false positives)
JUNK_EMAIL_DOMAINS = {
    "sentry.io", "wixpress.com", "example.com", "test.com",
    "domain.com", "email.com", "yoursite.com", "website.com",
    "company.com", "yourdomain.com",
}
JUNK_EMAIL_PREFIXES = {"noreply", "no-reply", "mailer-daemon", "postmaster"}


# ──────────────────────────────────────────────
# DEDUPLICATION STORE
# ──────────────────────────────────────────────
def load_seen(path: str) -> set:
    """Load previously scraped place keys from disk."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return set(json.load(f))
        except Exception:
            pass
    return set()


def save_seen(path: str, seen: set):
    """Persist seen place keys to disk."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(list(seen), f, indent=2)


def load_progress(path: str) -> dict:
    """Load query progress from disk. Returns dict mapping session_key -> list of completed queries."""
    if os.path.exists(path):
        try:
            with open(path, "r", encoding="utf-8") as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_progress(path: str, progress: dict):
    """Persist query progress to disk."""
    with open(path, "w", encoding="utf-8") as f:
        json.dump(progress, f, indent=2)


def make_session_key(query: str, location: str, with_websites: bool) -> str:
    """Create a unique key for this scraping session (query+location+mode)."""
    mode = "with_web" if with_websites else "no_web"
    return f"{query.strip().lower()}|{location.strip().lower()}|{mode}"


def place_key(href: str) -> str:
    """
    Extract a stable identifier from a Maps URL.
    Uses the /maps/place/<slug> portion so minor URL param differences
    between runs still resolve to the same key.
    """
    try:
        parsed = urllib.parse.urlparse(href)
        # path looks like /maps/place/Business+Name/@lat,lng,...
        parts = parsed.path.split("/")
        # ['', 'maps', 'place', 'Business+Name', '@...']
        if len(parts) >= 4:
            return urllib.parse.unquote_plus(parts[3]).lower().strip()
    except Exception:
        pass
    return href  # fallback: use full URL


# ──────────────────────────────────────────────
# HELPERS
# ──────────────────────────────────────────────
def clean_phone(raw: str) -> str:
    return re.sub(r"[^\d\+\-\(\) ]", "", raw).strip()


def is_junk_email(email: str) -> bool:
    """Filter out common false-positive emails."""
    email_lower = email.lower()
    local, _, domain = email_lower.partition("@")
    if domain in JUNK_EMAIL_DOMAINS:
        return True
    if local in JUNK_EMAIL_PREFIXES:
        return True
    # Filter image file extensions that regex might match
    if email_lower.endswith((".png", ".jpg", ".jpeg", ".gif", ".svg", ".webp", ".css", ".js")):
        return True
    return False


def safe_output_path(path: str) -> str:
    """
    If the target .xlsx file is locked (open in Excel), generate a
    timestamped filename instead of crashing with PermissionError.
    """
    if not path.endswith(".xlsx"):
        return path
    if not os.path.exists(path):
        return path
    try:
        # Quick write-test: try opening in append mode
        with open(path, "a"):
            pass
        return path
    except PermissionError:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        base, ext = os.path.splitext(path)
        new_path = f"{base}_{ts}{ext}"
        print(f"[!] '{path}' is open in another program.")
        print(f"[!] Saving to '{new_path}' instead.\n")
        return new_path


def build_query_url(query: str, location: str) -> str:
    """Combine query + optional location into a Google Maps search URL."""
    full_query = query.strip()
    if location:
        if full_query:
            full_query = f"{full_query} in {location.strip()}"
        else:
            full_query = location.strip()
    encoded = urllib.parse.quote_plus(full_query)
    return f"https://www.google.com/maps/search/{encoded}"


# ──────────────────────────────────────────────
# EMAIL & SOCIAL MEDIA EXTRACTION
# ──────────────────────────────────────────────
async def extract_email_from_maps_page(detail_page) -> str:
    """
    Try to find an email address directly on the Google Maps listing page.
    Some businesses list their email on their Maps profile.
    """
    emails = set()
    try:
        # Google Maps sometimes shows email in the info section
        # Check for mailto: links on the page
        mailto_links = await detail_page.locator("a[href^='mailto:']").all()
        for link in mailto_links:
            try:
                href = await link.get_attribute("href")
                if href:
                    email = href.replace("mailto:", "").split("?")[0].strip()
                    if email and not is_junk_email(email):
                        emails.add(email.lower())
            except Exception:
                pass

        # Also scan the page text for email patterns
        page_text = await detail_page.inner_text("body")
        for match in EMAIL_REGEX.findall(page_text):
            if not is_junk_email(match):
                emails.add(match.lower())
    except Exception:
        pass

    return "; ".join(sorted(emails))


async def search_email_for_business(context, business_name: str, location: str) -> str:
    """
    Try to find a business email by doing a quick Google search for
    "<business name> <location> email" and scanning the results.
    """
    emails = set()
    search_query = f"{business_name} {location} email contact".strip()
    search_url = f"https://www.google.com/search?q={urllib.parse.quote_plus(search_query)}"

    try:
        search_page = await context.new_page()
        await search_page.goto(search_url, wait_until="domcontentloaded", timeout=12_000)
        await asyncio.sleep(2)

        # Scan search results page for email addresses
        page_text = await search_page.inner_text("body")
        for match in EMAIL_REGEX.findall(page_text):
            if not is_junk_email(match):
                emails.add(match.lower())

        await search_page.close()
    except Exception:
        try:
            await search_page.close()
        except Exception:
            pass

    return "; ".join(sorted(emails))


async def extract_contact_info(context, website_url: str) -> dict:
    """
    Visit a business website and extract:
      - Email addresses (from mailto: links and raw text)
      - Social media links (Instagram, Facebook, Twitter, LinkedIn, TikTok, YouTube)

    Returns a dict like:
      {"Email": "info@biz.com", "Instagram": "https://instagram.com/biz", ...}
    """
    result = {
        "Email": "",
        "Instagram": "",
        "Facebook": "",
        "Twitter": "",
        "LinkedIn": "",
        "TikTok": "",
        "YouTube": "",
    }

    if not website_url:
        return result

    try:
        site_page = await context.new_page()
        await site_page.goto(website_url, wait_until="domcontentloaded", timeout=15_000)
        await asyncio.sleep(2)

        # Get full page HTML for scanning
        html = await site_page.content()

        # ── Extract emails ─────────────────────────────────────
        emails = set()

        # 1. mailto: links
        mailto_links = await site_page.locator("a[href^='mailto:']").all()
        for link in mailto_links:
            try:
                href = await link.get_attribute("href")
                if href:
                    email = href.replace("mailto:", "").split("?")[0].strip()
                    if email and not is_junk_email(email):
                        emails.add(email)
            except Exception:
                pass

        # 2. Regex scan on HTML
        for match in EMAIL_REGEX.findall(html):
            if not is_junk_email(match):
                emails.add(match.lower())

        if emails:
            result["Email"] = "; ".join(sorted(emails))

        # ── Extract social media links ─────────────────────────
        for platform, pattern in SOCIAL_PATTERNS.items():
            matches = pattern.findall(html)
            if matches:
                # Take the first clean match, strip trailing junk
                url = matches[0].rstrip("/").rstrip("\\")
                result[platform] = url

        # Also try checking a /contact or /about page for more info
        for sub_path in ["/contact", "/contact-us", "/about", "/about-us"]:
            try:
                contact_url = website_url.rstrip("/") + sub_path
                response = await site_page.goto(contact_url, wait_until="domcontentloaded", timeout=8_000)
                if response and response.ok:
                    await asyncio.sleep(1)
                    contact_html = await site_page.content()

                    # Emails from contact page
                    for link in await site_page.locator("a[href^='mailto:']").all():
                        try:
                            href = await link.get_attribute("href")
                            if href:
                                email = href.replace("mailto:", "").split("?")[0].strip()
                                if email and not is_junk_email(email):
                                    emails.add(email)
                        except Exception:
                            pass
                    for match in EMAIL_REGEX.findall(contact_html):
                        if not is_junk_email(match):
                            emails.add(match.lower())

                    # Social from contact page (only fill if not already found)
                    for platform, pattern in SOCIAL_PATTERNS.items():
                        if not result[platform]:
                            matches = pattern.findall(contact_html)
                            if matches:
                                result[platform] = matches[0].rstrip("/").rstrip("\\")

                    # Update email after contact page scan
                    if emails:
                        result["Email"] = "; ".join(sorted(emails))

                    break  # stop after first successful contact/about page
            except Exception:
                continue

        await site_page.close()

    except Exception as e:
        print(f"      [!] Could not scrape website ({type(e).__name__}): {website_url}")
        try:
            await site_page.close()
        except Exception:
            pass

    return result


# ──────────────────────────────────────────────
# SCRAPER
# ──────────────────────────────────────────────
async def scrape(query: str, location: str, max_results: int, output: str, with_websites: bool = False):
    seen: set = load_seen(SEEN_FILE)
    mode_label = "WITH websites" if with_websites else "WITHOUT websites"
    print(f"[*] Mode: collecting businesses {mode_label}")
    print(f"[*] Filter: skipping businesses WITHOUT a phone number")
    print(f"[*] Target: {max_results} matching businesses")
    print(f"[*] Loaded {len(seen)} previously seen businesses (will skip duplicates)")

    # Determine search queries
    if query:
        queries = [query]
    else:
        # No-niche mode: use broad search terms to discover diverse businesses
        queries = BROAD_QUERIES
        print(f"[*] No-niche mode: searching {len(queries)} broad categories")
        print(f"    Categories: {', '.join(queries)}")

    # ── Expand Location into Richest Cities (If Applicable) ───────────────
    actual_locations = [location]
    if location and location.lower() in RICHEST_CITIES:
        city_list = RICHEST_CITIES[location.lower()]
        actual_locations = [f"{city}, {location}" for city in city_list]
        print(f"[*] Country '{location}' detected. Searching richest cities first:")
        print(f"    {', '.join(city_list)}")

    # ── Create Query Configurations (Combinations of Query + Location) ────
    combinations = []
    for loc in actual_locations:
        for q in queries:
            combinations.append((q, loc))

    # ── Load query progress for resume ────────────────────────────────────
    progress = load_progress(PROGRESS_FILE)
    session_key = make_session_key(query, location, with_websites)
    completed_items = set(progress.get(session_key, []))

    # Filter out already-completed queries
    remaining_combos = []
    for q, loc in combinations:
        item_key = f"{q}||{loc}"
        if item_key in completed_items or (len(actual_locations) == 1 and q in completed_items):
            continue
        remaining_combos.append((q, loc))

    if len(remaining_combos) < len(combinations):
        skipped = len(combinations) - len(remaining_combos)
        print(f"[*] Resuming: skipping {skipped} already-completed queries/locations")
    if not remaining_combos:
        print(f"[*] All queries already completed for this session. Use --reset-progress to start over.")
        return

    results      = []
    new_seen_keys: set = set()

    async with async_playwright() as p:
        browser = await p.chromium.launch(headless=True, args=["--no-sandbox"])
        context = await browser.new_context(
            viewport={"width": 1280, "height": 900},
            locale="en-US",
        )

        scraped_count = 0
        total_inspected = 0

        for q_idx, (search_query, current_location) in enumerate(remaining_combos, 1):
            if scraped_count >= max_results:
                break

            page = await context.new_page()

            # ── 1. Navigate ──────────────────────────────────────────────────
            url = build_query_url(search_query, current_location)
            if len(remaining_combos) > 1:
                print(f"\n[*] Query {q_idx}/{len(remaining_combos)}: \"{search_query}\" in {current_location}")
            else:
                print(f"\n[*] Query: \"{search_query}\" in {current_location}")
            print(f"[*] Navigating to: {url}")
            await page.goto(url, wait_until="domcontentloaded", timeout=60_000)
            await asyncio.sleep(3)

            # ── 2. Accept cookies if prompted ────────────────────────────────
            try:
                btn = page.locator("button", has_text="Accept all")
                if await btn.count():
                    await btn.first.click()
                    await asyncio.sleep(1)
            except Exception:
                pass

            # ── 3. Scroll-collect-inspect loop ───────────────────────────────
            # Instead of collecting all URLs first and then filtering, we
            # continuously scroll for more listings whenever we need them.
            # This ensures --max N actually finds N matching businesses.
            PANEL = "div[role='feed']"
            print("[*] Scrolling and inspecting results …")

            collected_hrefs: list[str] = []  # ordered unique hrefs from this query
            collected_set: set = set()       # fast lookup for dedup
            inspect_cursor = 0               # next index in collected_hrefs to inspect
            stale_scroll_count = 0           # consecutive scrolls with no new links
            MAX_STALE_SCROLLS = 15           # give up on this query after this many

            while scraped_count < max_results:
                # ── Scroll for more listings if we've inspected everything so far
                if inspect_cursor >= len(collected_hrefs):
                    new_found_this_scroll = 0
                    # Do a few scroll rounds to collect more links
                    for _ in range(5):
                        links = await page.locator("a[href*='/maps/place/']").all()
                        for link in links:
                            href = await link.get_attribute("href")
                            if href and href not in collected_set:
                                collected_set.add(href)
                                collected_hrefs.append(href)
                                new_found_this_scroll += 1

                        await page.evaluate(
                            f'const el = document.querySelector("{PANEL}"); if (el) el.scrollTop += 800;'
                        )
                        await asyncio.sleep(SCROLL_PAUSE)

                    if new_found_this_scroll == 0:
                        stale_scroll_count += 1
                        if stale_scroll_count >= MAX_STALE_SCROLLS:
                            print(f"[*] No more results to scroll for this query.")
                            break
                    else:
                        stale_scroll_count = 0

                    # Still nothing new to inspect? Keep scrolling
                    if inspect_cursor >= len(collected_hrefs):
                        continue

                # ── Pick next listing to inspect
                href = collected_hrefs[inspect_cursor]
                inspect_cursor += 1
                total_inspected += 1

                key = place_key(href)

                # Dedup: skip already-seen across queries in this run
                if key in new_seen_keys:
                    continue

                # Cross-run duplicate check
                if key in seen:
                    print(f"  [{total_inspected}] SKIP (already scraped in a previous run): {key}")
                    continue

                try:
                    detail_page = await context.new_page()
                    await detail_page.goto(href, wait_until="domcontentloaded", timeout=60_000)
                    await asyncio.sleep(2)

                    # ── Name ──────────────────────────────────────────────
                    name = ""
                    try:
                        name = await detail_page.locator("h1").first.inner_text(timeout=5_000)
                    except Exception:
                        pass

                    # ── Website & Menu Fallback ───────────────────────────
                    website = ""
                    try:
                        web_link = detail_page.locator("a[data-item-id='authority']")
                        if await web_link.count():
                            website = await web_link.first.get_attribute("href") or ""
                            
                        # If no website is found, check the menu link (often serves as their website)
                        if not website:
                            menu_link = detail_page.locator("a[data-item-id='menu']")
                            if await menu_link.count():
                                website = await menu_link.first.get_attribute("href") or ""
                    except Exception:
                        pass

                    if with_websites:
                        # We WANT businesses that have a website — skip those without
                        if not website:
                            new_seen_keys.add(key)
                            await detail_page.close()
                            print(f"  [{total_inspected}] SKIP (no website): {name}")
                            continue
                    else:
                        # We WANT businesses WITHOUT a website — skip those with one
                        if website:
                            new_seen_keys.add(key)
                            await detail_page.close()
                            print(f"  [{total_inspected}] SKIP (has website): {name}")
                            continue

                    # ── Phone ─────────────────────────────────────────────
                    phone = ""
                    try:
                        phone_el = detail_page.locator("button[data-item-id*='phone']")
                        if await phone_el.count():
                            phone = clean_phone(await phone_el.first.get_attribute("aria-label") or "")
                            phone = phone.replace("Phone:", "").strip()
                    except Exception:
                        pass

                    # Skip businesses without a phone number
                    if not phone:
                        new_seen_keys.add(key)
                        await detail_page.close()
                        print(f"  [{total_inspected}] SKIP (no phone number): {name}")
                        continue

                    # ── Address ───────────────────────────────────────────
                    address = ""
                    try:
                        addr_el = detail_page.locator("button[data-item-id='address']")
                        if await addr_el.count():
                            address = await addr_el.first.get_attribute("aria-label") or ""
                            address = address.replace("Address:", "").strip()
                    except Exception:
                        pass

                    # ── Rating & Reviews ──────────────────────────────────
                    rating, reviews = "", ""
                    try:
                        rating_el = detail_page.locator("div.F7nice span[aria-hidden='true']")
                        if await rating_el.count():
                            rating = await rating_el.first.inner_text()
                        review_el = detail_page.locator("div.F7nice span[aria-label*='reviews']")
                        if await review_el.count():
                            reviews = (await review_el.first.get_attribute("aria-label") or "").split()[0]
                    except Exception:
                        pass

                    # ── Category ──────────────────────────────────────────
                    category = ""
                    try:
                        cat_el = detail_page.locator("button.DkEaL")
                        if await cat_el.count():
                            category = await cat_el.first.inner_text()
                    except Exception:
                        pass

                    # ── Email & Social Media ──────────────────────────────
                    contact_info = {
                        "Email": "", "Instagram": "", "Facebook": "",
                        "Twitter": "", "LinkedIn": "", "TikTok": "", "YouTube": "",
                    }

                    if website:
                        # Has website → scrape email + socials from it
                        await detail_page.close()
                        print(f"      Scraping website for email/socials: {website[:60]}…")
                        contact_info = await extract_contact_info(context, website)
                    else:
                        # No website → try to find email from Maps page + Google search
                        print(f"      Searching for email (no website)…")
                        maps_email = await extract_email_from_maps_page(detail_page)
                        await detail_page.close()
                        if maps_email:
                            contact_info["Email"] = maps_email
                        else:
                            # Try a quick Google search for the business email
                            search_loc = current_location if current_location else ""
                            google_email = await search_email_for_business(context, name, search_loc)
                            if google_email:
                                contact_info["Email"] = google_email

                    extras = []
                    if contact_info["Email"]:
                        extras.append(f"📧 {contact_info['Email'][:40]}")
                    social_found = [p for p in SOCIAL_PATTERNS if contact_info.get(p)]
                    if social_found:
                        extras.append(f"🔗 {', '.join(social_found)}")
                    if extras:
                        print(f"      Found: {' | '.join(extras)}")

                    record = {
                        "Name":      name.strip(),
                        "Category":  category.strip(),
                        "Phone":     phone,
                        "Email":     contact_info["Email"],
                        "Address":   address,
                        "Rating":    rating.strip(),
                        "Reviews":   reviews.strip(),
                        "Website":   website,
                        "Instagram": contact_info["Instagram"],
                        "Facebook":  contact_info["Facebook"],
                        "Twitter":   contact_info["Twitter"],
                        "LinkedIn":  contact_info["LinkedIn"],
                        "TikTok":    contact_info["TikTok"],
                        "YouTube":   contact_info["YouTube"],
                        "Maps URL":  href,
                    }

                    results.append(record)
                    new_seen_keys.add(key)
                    scraped_count += 1
                    remaining = max_results - scraped_count
                    print(f"  [{total_inspected}] ✓ ({scraped_count}/{max_results}) {name} | {phone} | {address[:45]}")

                except Exception as e:
                    print(f"  [{total_inspected}] ERROR: {e}")
                finally:
                    try:
                        await detail_page.close()
                    except Exception:
                        pass

            await page.close()

            # Mark this query config as completed in progress
            item_key = f"{search_query}||{current_location}"
            completed_items.add(item_key)
            progress[session_key] = list(completed_items)
            save_progress(PROGRESS_FILE, progress)

            print(f"[*] Progress: {scraped_count}/{max_results} matching businesses found so far")
            print(f"[*] Query \"{search_query}\" in {current_location} completed and saved to progress")

        await browser.close()

    # ── 5. Persist seen keys ──────────────────────────────────────────────
    seen.update(new_seen_keys)
    save_seen(SEEN_FILE, seen)
    print(f"\n[*] Updated seen list → {len(seen)} total businesses tracked in {SEEN_FILE}")

    # ── 6. Save results ───────────────────────────────────────────────────
    if not results:
        print("\n[!] No new matching businesses found.")
        return

    df = pd.DataFrame(results)
    df.index = range(1, len(df) + 1)

    # Resolve output path (handles locked file)
    output = safe_output_path(output)

    sheet_name = "Website Leads" if with_websites else "No Website Leads"

    if output.endswith(".xlsx"):
        try:
            with pd.ExcelWriter(output, engine="openpyxl") as writer:
                df.to_excel(writer, sheet_name=sheet_name, index_label="#")
                ws = writer.sheets[sheet_name]
                for col in ws.columns:
                    max_len = max(len(str(cell.value or "")) for cell in col) + 4
                    ws.column_dimensions[col[0].column_letter].width = min(max_len, 60)
            print(f"\n[✓] Saved {len(df)} leads → {output}")
        except PermissionError:
            # Last-resort fallback to CSV
            csv_path = output.replace(".xlsx", ".csv")
            df.to_csv(csv_path, index_label="#")
            print(f"\n[!] Could not write Excel file. Saved as CSV → {csv_path}")
    else:
        df.to_csv(output, index_label="#")
        print(f"\n[✓] Saved {len(df)} leads → {output}")

    # ── 7. Summary ────────────────────────────────────────────────────────
    print(f"\n{'─'*55}")
    print(f"  Total leads collected : {len(df)}")
    print(f"  With phone number     : {df['Phone'].astype(bool).sum()}")
    print(f"  With email            : {df['Email'].astype(bool).sum()}")
    print(f"  With address          : {df['Address'].astype(bool).sum()}")
    if 'Website' in df.columns:
        print(f"  With website          : {df['Website'].astype(bool).sum()}")
    social_cols = ["Instagram", "Facebook", "Twitter", "LinkedIn", "TikTok", "YouTube"]
    for col in social_cols:
        if col in df.columns:
            count = df[col].astype(bool).sum()
            if count:
                print(f"  With {col:17s} : {count}")
    print(f"{'─'*55}")


# ──────────────────────────────────────────────
# ENTRY POINT
# ──────────────────────────────────────────────
def main():
    parser = argparse.ArgumentParser(
        description="Scrape Google Maps for businesses — with email & social media extraction.",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""\
Examples:
  # With a niche + location (city, region, or country)
  python gmaps_scraper.py --query "restaurants" --location "India"
  python gmaps_scraper.py --query "plumbers" --location "Chicago IL" --max 50

  # Country/region level (no specific city needed)
  python gmaps_scraper.py --query "restaurants" --location "UK" --max 50

  # With websites + email/social scraping
  python gmaps_scraper.py --query "dentists" --location "Kosovo" --with-websites
  python gmaps_scraper.py --query "hotels" --location "Tokyo" --with-websites --output tokyo_leads.xlsx

  # Location only — no niche (discovers diverse businesses)
  python gmaps_scraper.py --location "UK" --max 50
  python gmaps_scraper.py --location "Berlin" --with-websites --max 30
        """
    )
    parser.add_argument("--query",    default="",
                        help='Business type, e.g. "restaurants" or "dentists". '
                             'Optional — omit to scrape diverse business types.')
    parser.add_argument("--location", default="",
                        help='Location: country, region, or city, e.g. "UK", "India", "Austin TX". '
                             'Works without a city — country/region level is fine.')
    parser.add_argument("--max",      type=int, default=DEFAULT_MAX,
                        help=f"Max leads to collect (default {DEFAULT_MAX})")
    parser.add_argument("--output",   default=OUTPUT_FILE,
                        help=f"Output file (.xlsx or .csv). Default: {OUTPUT_FILE}")
    parser.add_argument("--with-websites", action="store_true",
                        help="Collect businesses that DO have a website (default: collect those WITHOUT). "
                             "Also scrapes emails & social media links from the website.")
    parser.add_argument("--reset-seen", action="store_true",
                        help=f"Clear the seen-businesses cache ({SEEN_FILE}) and start fresh")
    parser.add_argument("--reset-progress", action="store_true",
                        help=f"Clear the query progress cache ({PROGRESS_FILE}) so all queries run again")
    args = parser.parse_args()

    if not args.query and not args.location:
        parser.error("At least one of --query or --location is required.")

    if args.reset_seen and os.path.exists(SEEN_FILE):
        os.remove(SEEN_FILE)
        print(f"[*] Cleared seen cache ({SEEN_FILE})\n")

    if args.reset_progress and os.path.exists(PROGRESS_FILE):
        os.remove(PROGRESS_FILE)
        print(f"[*] Cleared query progress cache ({PROGRESS_FILE})\n")

    asyncio.run(scrape(args.query, args.location, args.max, args.output, args.with_websites))


if __name__ == "__main__":
    main()
