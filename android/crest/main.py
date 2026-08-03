# ============================================================
# main.py — boAt Crest App Review Analyzer (Cloud Version)
# Author  : boAt R&D (QA)
# Replaces: CrestApp_WebScraping, All_HtmlFiles_Csv,
#           App_Rating_Finder, Csv_Date_Rating_Filter_7days,
#           CSV_Remove_Duplicates, Csv_2_Xlsx
# ============================================================

import os
import json
import uuid
import time
import pytz
import openpyxl
import pandas as pd
import BFAT_CrestApp_PlayStore_Code as orig
from datetime import datetime, timedelta
from google.oauth2 import service_account
from googleapiclient.discovery import build
from google_play_scraper import app as gplay_app

from BFAT_CrestApp_PlayStore_Code import (
    SheetNameChange,
    Positive_Reviews,
    Negative_Reviews,
    Sorting_Xlsx,
    Sorting_Summary_Xlsx,
    Negative_Review_Analysis,
    Review_Date_Calculator,
    Summary,
    Summary_Enhance,
    close_excel_file,
)

# ── CONFIG ──────────────────────────────────────────────────
APP_ID    = os.environ.get('ANDROID_APP_ID')
if not APP_ID:
    raise SystemExit("❌ Missing required secret: ANDROID_APP_ID")
XLSX_FILE = 'Android_CrestApp_Review.xlsx'
SCOPES    = ['https://www.googleapis.com/auth/androidpublisher']

# ── GLOBALS ─────────────────────────────────────────────────
App_Rating  = None  # calculated by App_Rating_Finder() from this week's reviews
App_Version = "N/A"
lowest_date = None
latest_date = None

# ── SDK LEVEL → ANDROID VERSION NAME MAPPING ────────────────
ANDROID_SDK_MAP = {
    36: "16", 35: "15", 34: "14", 33: "13", 32: "12L", 31: "12",
    30: "11", 29: "10", 28: "9", 27: "8.1", 26: "8.0", 25: "7.1",
    24: "7.0", 23: "6.0", 22: "5.1", 21: "5.0", 20: "4.4W",
    19: "4.4", 18: "4.3", 17: "4.2", 16: "4.1", 15: "4.0.3",
    14: "4.0",
}

def sdk_to_android_version(sdk):
    if not sdk:
        return ''
    try:
        sdk_int = int(sdk)
    except (ValueError, TypeError):
        return str(sdk)
    return ANDROID_SDK_MAP.get(sdk_int, f"SDK {sdk_int}")

# ── GOOGLE PLAY DEVELOPER API SERVICE ───────────────────────
def get_play_service():
    creds_json = os.environ.get('GOOGLE_PLAY_CREDENTIALS_JSON')
    if creds_json:
        creds_info = json.loads(creds_json)
        creds = service_account.Credentials.from_service_account_info(
            creds_info, scopes=SCOPES
        )
    else:
        # Local testing fallback — reads from a local json key file
        creds = service_account.Credentials.from_service_account_file(
            'service_account.json', scopes=SCOPES
        )
    return build('androidpublisher', 'v3', credentials=creds)

# ── STEP 1: FETCH REVIEWS ───────────────────────────────────
def fetch_and_save_reviews():
    print("=" * 60)
    print("Step 1: Fetching reviews from Google Play Developer API...")
    print("=" * 60)

    service = get_play_service()

    all_reviews  = []
    token        = None

    while True:
        request = service.reviews().list(
            packageName=APP_ID,
            maxResults=100,
            token=token,
            translationLanguage='en',
        )
        response = request.execute()
        batch = response.get('reviews', [])
        all_reviews.extend(batch)

        token = response.get('tokenPagination', {}).get('nextPageToken')
        if not token or not batch:
            break

    print(f"  Raw reviews fetched from API: {len(all_reviews)}")

    rows = []
    for r in all_reviews:
        review_id = r.get('reviewId', str(uuid.uuid4()))
        author    = r.get('authorName', 'Anonymous')

        comments = r.get('comments', [])
        if not comments:
            continue
        user_comment = comments[0].get('userComment', {})

        # Use translated text (English) — 'text' field returns translated
        # version when translationLanguage param is set on the request.
        text = user_comment.get('text', '') or user_comment.get('originalText', '')

        star_rating     = user_comment.get('starRating', '')
        app_version     = user_comment.get('appVersionName', '')
        app_version_code = user_comment.get('appVersionCode', '')
        android_os_ver  = sdk_to_android_version(user_comment.get('androidOsVersion', ''))
        reviewer_lang   = user_comment.get('reviewerLanguage', '')

        device_meta  = user_comment.get('deviceMetadata', {})
        phone_name   = device_meta.get('productName', '') or user_comment.get('device', '')

        # Timestamp comes as epoch seconds (string) inside lastModified
        last_modified = user_comment.get('lastModified', {})
        seconds       = int(last_modified.get('seconds', 0))
        review_dt     = datetime.utcfromtimestamp(seconds) if seconds else None

        # Developer reply, if any
        reply_text = ''
        if len(comments) > 1:
            dev_comment = comments[1].get('developerComment', {})
            reply_text  = dev_comment.get('text', '')

        rows.append({
            'reviewId':      review_id,
            'User Name':     author,
            'Ratings':       star_rating,
            'Reviews':       text,
            'Phone Name':    phone_name,
            'OS Version':    android_os_ver,
            'Date Time':     review_dt,
            'Version Code':  app_version_code,
            'App Version':   app_version,
            'boAt Reply':    reply_text,
        })

    df = pd.DataFrame(rows)

    if df.empty:
        print("  No reviews returned by API.")
        df.to_excel(XLSX_FILE, index=False)
        return

    # ── Use IST timezone so GitHub Actions (UTC) and local machine
    #    both use the same cutoff.
    ist = pytz.timezone('Asia/Kolkata')
    now_ist = datetime.now(ist).replace(tzinfo=None)
    seven_days_ago = now_ist - timedelta(days=7)

    # API timestamps are in UTC — convert to IST for correct comparison
    df['Date Time'] = pd.to_datetime(df['Date Time'])
    df['Date Time'] = df['Date Time'].dt.tz_localize('UTC').dt.tz_convert(ist).dt.tz_localize(None)
    df = df[df['Date Time'] >= seven_days_ago]

    print(f"  IST now        : {now_ist.strftime('%Y-%m-%d %H:%M:%S')} IST")
    print(f"  Cutoff (7 days): {seven_days_ago.strftime('%Y-%m-%d %H:%M:%S')} IST")

    # Reorder columns to match original manual scraper format exactly
    df = df[[
        'reviewId', 'User Name', 'Ratings', 'Reviews',
        'Phone Name', 'OS Version', 'Date Time',
        'Version Code', 'App Version', 'boAt Reply'
    ]]

    # Sort and deduplicate
    df = df.sort_values('Date Time', ascending=False)
    df = df.drop_duplicates(subset=['User Name', 'Ratings', 'Reviews'])

    # Save to Excel
    df.to_excel(XLSX_FILE, index=False)

    print(f"  Fetched : {len(df)} reviews from last 7 days")
    print(f"  Saved   : {XLSX_FILE}")

# ── STEP 2: GET APP VERSION ─────────────────────────────────
def App_Version_Finder():
    global App_Version

    workbook = openpyxl.load_workbook(XLSX_FILE)
    sheet    = workbook.active

    # App Version is column I (9) per the new column order
    APP_VERSION_COL = 9
    highest_version = None

    for row in range(2, sheet.max_row + 1):
        ver = sheet.cell(row=row, column=APP_VERSION_COL).value
        if ver:
            ver_str = str(ver).split(":")[-1].strip()
            if highest_version is None or ver_str > highest_version:
                highest_version = ver_str

    App_Version = highest_version if highest_version else "N/A"
    print(f"  Highest App Version: {App_Version}")
    workbook.close()

# ── STEP 2b: FETCH LIVE APP RATING + VERSION (public Play Store scrape) ─────
def App_Rating_Finder():
    """
    Google Play Developer API (androidpublisher) has no endpoint for the
    app's overall public star rating — that figure only exists on the
    public Play Store listing page, not through any authenticated API.

    Uses google-play-scraper (unofficial, no auth) to read the same
    'score' and 'version' values shown live on the Play Store page —
    the real number customers see, not an approximation.

    Falls back to averaging this week's fetched reviews (old behaviour)
    only if the live scrape fails for any reason, so the pipeline never
    breaks or reports a blank N/A when a fallback is possible.
    """
    global App_Rating, App_Version

    try:
        result = gplay_app(APP_ID, lang='en', country='in')
        score   = result.get('score')
        version = result.get('version')

        App_Rating  = f"{score:.1f}" if isinstance(score, (int, float)) else "N/A"
        if version:
            App_Version = str(version)

        print(f"  Live Play Store Rating  : {App_Rating}")
        print(f"  Live Play Store Version : {App_Version}")
        return
    except Exception as e:
        print(f"  ⚠️  Live Play Store scrape failed ({e}); "
              f"falling back to this week's review average.")

    # ── Fallback: average of this week's fetched review ratings ────────────
    workbook = openpyxl.load_workbook(XLSX_FILE)
    sheet    = workbook.active

    RATINGS_COL = 3  # column C — Ratings
    ratings = []

    for row in range(2, sheet.max_row + 1):
        val = sheet.cell(row=row, column=RATINGS_COL).value
        if val is not None:
            try:
                ratings.append(float(val))
            except (ValueError, TypeError):
                continue

    if ratings:
        avg = sum(ratings) / len(ratings)
        App_Rating = f"{avg:.1f}"
    else:
        App_Rating = "N/A"

    print(f"  Fallback Average App Rating (this week's reviews): {App_Rating}")
    workbook.close()

# ── MAIN ────────────────────────────────────────────────────
if __name__ == '__main__':

    # Step 1 — Fetch reviews from Play Store
    fetch_and_save_reviews()
    time.sleep(2)

    # Step 2 — Rename Sheet1 → CrestApp_Review_Data
    SheetNameChange()

    # Step 3 — Split into Positive / Negative sheets
    Positive_Reviews()
    Negative_Reviews()

    # Step 4 — Format all sheets
    Sorting_Xlsx('Positive_Reviews')
    Sorting_Xlsx('Negative_Reviews')
    Sorting_Xlsx('CrestApp_Review_Data')

    # Step 5 — Analysis
    Negative_Review_Analysis()

    # Step 6 — Get App Version + Date Range
    App_Version_Finder()
    App_Rating_Finder()
    Review_Date_Calculator()

    # Step 7 — Inject App Version / Rating into the BFAT module before Summary
    # (lowest_date / latest_date are already set correctly inside orig by
    #  Review_Date_Calculator() above — do NOT overwrite them here)
    orig.App_Version = App_Version
    orig.App_Rating  = App_Rating

    # Step 8 — Generate Summary sheet
    Summary()
    time.sleep(3)

    # Step 9 — Format and enhance Summary
    Sorting_Summary_Xlsx('Summary')
    time.sleep(3)
    Summary_Enhance()
    time.sleep(3)

    # Step 10 — Close file
    close_excel_file(XLSX_FILE)

    print("=" * 60)
    print("Pipeline complete! Android_CrestApp_Review.xlsx ready.")
    print("=" * 60)
