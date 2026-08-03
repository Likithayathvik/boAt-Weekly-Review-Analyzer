"""
main.py — iOS Wearables App Review Fetcher
Fetches reviews from App Store Connect API and saves to Excel
"""

import re
import os
import jwt
import time
import requests
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment
from datetime import datetime, timedelta, timezone

# ── CONFIG (all values from GitHub Secrets, no hardcoded fallback) ──────────
KEY_ID    = os.environ.get('APPLE_KEY_ID')
ISSUER_ID = os.environ.get('APPLE_ISSUER_ID')
APP_ID    = os.environ.get('APPLE_APP_ID')
XLSX_FILE = "iOS_WearablesApp_Review.xlsx"
DAYS_BACK = 7
APP_NAME  = os.environ.get('APP_DISPLAY_NAME', "iOS Wearables App")

if not all([KEY_ID, ISSUER_ID, APP_ID]):
    raise SystemExit(
        "❌ Missing required secrets: APPLE_KEY_ID, APPLE_ISSUER_ID, APPLE_APP_ID "
        "must be set as environment variables / GitHub Secrets."
    )

PRIVATE_KEY = os.environ.get('APPLE_PRIVATE_KEY')
if not PRIVATE_KEY:
    raise SystemExit("❌ Missing required secret: APPLE_PRIVATE_KEY")

# ── KEYWORD PATTERNS ─────────────────────────────────────────────────────────
KEYWORD_PATTERNS = {
    'Connectivity':          r'\b(connect(ion|ing)?|not\s*connected|not\s*getting\s*connect(ed|ion)?|unable\s*to\s*connect|bad\s*connectivity)\b',
    'Device disconnection':  r'\bdisconnected?\b',
    'Login':                 r'\b(login|log(in|ged)?|otp|sign\s*in|signing\s*in)\b',
    'Location Permission':   r'\b(location|gps|geolocation)\s*(permission|access)?\b',
    'Phone becoming silent': r'\bsilent\b',
    'SKU Not supported':     r"(?:can't\s*connect.*app|not\s*(?:connect(?:ing)?|support(?:ed)?|show(?:n|ing)|found|recognized|compatible|listed|list|option)\s*(?:by\s*this\s*app)?|this\s*app\s*doesn't\s*support\s*all\s*boat\s*smartwatch|watch\s*doesn't\s*pair\s*and\s*show\s*in\s*the\s*app|boat\s*watch\s*not\s*listed|devices\s*could\s*not\s*found|not\s*available\s*in\s*list|supported\s*device\s*list|this\s*model\s*is\s*not\s*(?:listed|shown|supported|found)\s*in\s*the\s*app|cannot\s*pair\s*it\s*with\s*this\s*app|not\s*listed\s*in\s*app|not\s*list|no\s*option|option\s*is\s*not\s*available\s*in\s*this\s*app|can't\s*see|didn'?t\s*find\s*my\s*device|does\s*not\s*show\s*other\s*models|unable\s*to\s*find\s*my\s*product|didn'?t\s*found\s*my\s*device|can't\s*found\s*my\s*device|no\s*support\s*for\s*old\s*devices|not\s*supported\s*for\s*old\s*devices)",
    'Positive Feedback':     r'^(?!.*(not|never|no)).*\b(good|super|nice|nyc|best|beautiful|great|amazing|awesome|impressed)\b',
    'Powering off (Device)': r'\b(auto\s*on\s*off|powering\s*off|auto\s*shutdown)\b',
    'Crash':                 r'\b(crash(ing)?|hang(ing)?|freeze(ing)?|carshed)\b',
    'Package box (Device)':  r'\b(open\s*box|package\s*box)\b',
    'Battery':               r'\b(battery|battery\s*percentage|battery\s*level|case\s*battery|watch\s*battery)\b',
    'Watch Face':            r'\b(watch\s*faces?|cloud\s*watch\s*face|clock\s*face|smartwatch\s*faces?|watch\s*display|custom\s*watch\s*faces?)\b',
    'Firmware':              r'\b(firmware|firm\s*ware)\b',
}
COMPILED_PATTERNS = {k: re.compile(v, re.IGNORECASE) for k, v in KEYWORD_PATTERNS.items()}

# ── CLASSIFY REVIEW ──────────────────────────────────────────────────────────
def classify_review(review):
    if not isinstance(review, str): review = ""
    found_reasons = [k for k, p in COMPILED_PATTERNS.items() if p.search(review)]
    reason = ', '.join(found_reasons) if found_reasons else 'Generic Comments'
    if len(found_reasons) > 1:
        consolidated = 'Decision yet to be confirmed'
    else:
        consolidated = reason
    return reason, consolidated

# ── GENERATE JWT TOKEN ───────────────────────────────────────────────────────
def generate_token():
    payload = {
        "iss": ISSUER_ID,
        "iat": int(time.time()),
        "exp": int(time.time()) + 1200,
        "aud": "appstoreconnect-v1"
    }
    return jwt.encode(payload, PRIVATE_KEY, algorithm="ES256", headers={"kid": KEY_ID})

# ── FETCH REVIEWS ────────────────────────────────────────────────────────────
def fetch_reviews():
    print("=" * 60)
    print("Fetching iOS Wearables App reviews from App Store Connect API...")
    print("=" * 60)

    today     = datetime.now(timezone.utc)
    cutoff_dt = today - timedelta(days=DAYS_BACK)
    print(f"  Date range : {cutoff_dt.strftime('%d %b %Y')} → {today.strftime('%d %b %Y')}")

    token   = generate_token()
    headers = {"Authorization": f"Bearer {token}"}
    url     = f"https://api.appstoreconnect.apple.com/v1/apps/{APP_ID}/customerReviews"
    params  = {
        "sort":  "-createdDate",
        "limit": 200,
        "fields[customerReviews]": "title,body,rating,createdDate,reviewerNickname,territory,response"
    }

    all_reviews = []
    page = 1
    stop = False

    while url and not stop:
        print(f"  Fetching page {page}...")
        resp = requests.get(url, headers=headers, params=params)
        if resp.status_code == 401:
            raise SystemExit("❌ 401 Unauthorized — check Key ID / Issuer ID / Private Key")
        if resp.status_code != 200:
            raise SystemExit(f"❌ Error {resp.status_code}: {resp.text[:300]}")

        data  = resp.json()
        items = data.get("data", [])

        for item in items:
            created_str = item.get("attributes", {}).get("createdDate", "")
            try:
                created_dt = datetime.strptime(created_str[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
            except Exception:
                created_dt = today
            if created_dt < cutoff_dt:
                stop = True
                break
            all_reviews.append(item)

        url    = data.get("links", {}).get("next")
        params = {}
        page  += 1
        if page > 20:
            break

    print(f"  Total in date range : {len(all_reviews)} reviews")
    return all_reviews, cutoff_dt, today

# ── PARSE + CLASSIFY ─────────────────────────────────────────────────────────
def parse_reviews(raw):
    rows = []
    for item in raw:
        attr        = item.get("attributes", {})
        created_str = attr.get("createdDate", "")
        try:
            created_dt = datetime.strptime(created_str[:19], "%Y-%m-%dT%H:%M:%S").replace(tzinfo=timezone.utc)
        except Exception:
            created_dt = datetime.now(timezone.utc)

        rating   = attr.get("rating", 0)
        title    = attr.get("title",            "") or ""
        body     = attr.get("body",             "") or ""
        reviewer = attr.get("reviewerNickname", "") or "Anonymous"

        reason, consolidated = classify_review(body) if rating <= 3 else ("", "")

        rows.append({
            "reviewId":            item.get("id", ""),
            "User":                reviewer,
            "Rating":              rating,
            "Review Title":        title,
            "Review":              body,
            "Date":                created_dt.strftime("%Y-%m-%d"),
            "Reason":              reason,
            "Consolidated Reason": consolidated,
        })

    if rows:
        dates = [r["Date"] for r in rows]
        print(f"  Oldest : {min(dates)}   Newest : {max(dates)}")
    return rows

# ── STYLE HELPERS ────────────────────────────────────────────────────────────
DATA_COL_WIDTHS    = {'A': 15.0, 'C': 8.0, 'D': 28.0, 'E': 42.0, 'F': 11.0}
SUMMARY_COL_WIDTHS = {'A': 25.0, 'B': 15.0, 'C': 21.0, 'D': 23.0}

def style_data_sheet(ws):
    bold_font   = Font(bold=True,  name="Calibri", size=11)
    normal_font = Font(bold=False, name="Calibri", size=11)
    for cell in ws[1]:
        cell.font      = bold_font
        cell.alignment = Alignment(vertical="center", wrap_text=True)
    ws.row_dimensions[1].height = 43.5
    for row in ws.iter_rows(min_row=2):
        for cell in row:
            cell.font      = normal_font
            cell.alignment = Alignment(wrap_text=True)
    for col_letter, width in DATA_COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width

def style_summary_sheet(ws):
    bold_font   = Font(bold=True,  name="Calibri", size=11)
    normal_font = Font(bold=False, name="Calibri", size=11)
    center      = Alignment(horizontal="center", vertical="center")

    def fill(hex_color):
        return PatternFill("solid", fgColor=hex_color)

    for col in range(1, 5):
        cell = ws.cell(row=1, column=col)
        cell.font      = bold_font
        cell.fill      = fill("B9D9EB")
        cell.alignment = center if col > 1 else Alignment(vertical="center")

    for col in range(1, 5):
        cell = ws.cell(row=2, column=col)
        cell.font      = normal_font
        cell.alignment = center if col > 1 else Alignment()

    for col in range(1, 3):
        ws.cell(row=4, column=col).font = bold_font
        ws.cell(row=4, column=col).fill = fill("C6EFCE")
        ws.cell(row=4, column=col).alignment = center if col == 2 else Alignment()

    for col in range(1, 3):
        ws.cell(row=5, column=col).font = bold_font
        ws.cell(row=5, column=col).fill = fill("66CC66")
        ws.cell(row=5, column=col).alignment = center if col == 2 else Alignment()

    for col in range(1, 3):
        ws.cell(row=6, column=col).font = bold_font
        ws.cell(row=6, column=col).fill = fill("FFFF00")
        ws.cell(row=6, column=col).alignment = center if col == 2 else Alignment()

    for col in range(1, 3):
        ws.cell(row=8, column=col).font = bold_font
        ws.cell(row=8, column=col).fill = fill("FFF2CC")
        ws.cell(row=8, column=col).alignment = center if col == 2 else Alignment()

    for row_idx in range(9, ws.max_row + 1):
        ws.cell(row=row_idx, column=1).font = normal_font
        ws.cell(row=row_idx, column=2).font = normal_font
        ws.cell(row=row_idx, column=2).alignment = center

    for col_letter, width in SUMMARY_COL_WIDTHS.items():
        ws.column_dimensions[col_letter].width = width


# ── FETCH APP RATING + LATEST VERSION FROM ITUNES ────────────────────────────
def fetch_itunes_info():
    """Fetch overall app rating and latest version from iTunes public lookup API."""
    try:
        url  = f"https://itunes.apple.com/lookup?id={APP_ID}&country=in"
        resp = requests.get(url, timeout=10).json()
        results = resp.get("results", [])
        if results:
            r = results[0]
            rating  = round(r.get("averageUserRating", 0), 1)
            version = r.get("version", "N/A")
            print(f"  iTunes lookup: rating={rating}, version={version}")
            return str(rating), version
    except Exception as e:
        print(f"  ⚠️  iTunes lookup failed: {e}")
    return "N/A", "N/A"

# ── SAVE TO EXCEL ─────────────────────────────────────────────────────────────
def save_to_xlsx(rows, cutoff_dt, today):
    if not rows:
        print("\n⚠️  No reviews found in the last 7 days.")
        return

    df      = pd.DataFrame(rows)
    df_pos  = df[df["Rating"] >= 4][["reviewId","User","Rating","Review Title","Review","Date"]].copy()
    df_neg  = df[df["Rating"] <= 3][["reviewId","User","Rating","Review Title","Review","Date","Reason","Consolidated Reason"]].copy()
    df_main = df[["reviewId","User","Rating","Review Title","Review","Date"]].copy()

    date_range = f"{cutoff_dt.strftime('%Y-%m-%d')} to {today.strftime('%Y-%m-%d')}"

    with pd.ExcelWriter(XLSX_FILE, engine="openpyxl") as writer:
        df_main.to_excel(writer, sheet_name="Wapp_Review_Data",  index=False)
        df_pos.to_excel(writer,  sheet_name="Positive_Reviews",  index=False)
        df_neg.to_excel(writer,  sheet_name="Negative_Reviews",  index=False)

        wb     = writer.book
        ws_sum = wb.create_sheet("Summary")
        ws_sum.append(["App Name", "App Rating", "Latest App Version", "Review Data"])
        app_rating, app_version = fetch_itunes_info()
        ws_sum.append([APP_NAME, f"{app_rating}*", app_version, date_range])
        ws_sum.append([])
        ws_sum.append(["Category", "Count"])
        ws_sum.append(["Positive Reviews (4 & 5)", len(df_pos)])
        ws_sum.append(["Negative Reviews (<=3)",   len(df_neg)])
        ws_sum.append([])
        ws_sum.append(["Summary (-ve reviews)", "Count"])
        reason_counts = df_neg["Consolidated Reason"].value_counts() if len(df_neg) > 0 else {}
        for reason, count in reason_counts.items():
            ws_sum.append([reason, int(count)])
        style_summary_sheet(ws_sum)

    wb = openpyxl.load_workbook(XLSX_FILE)
    for sname in ["Wapp_Review_Data", "Positive_Reviews", "Negative_Reviews"]:
        if sname in wb.sheetnames:
            style_data_sheet(wb[sname])
    style_summary_sheet(wb["Summary"])
    wb.save(XLSX_FILE)

    print(f"\n  Total    : {len(df)}")
    print(f"  Positive : {len(df_pos)}")
    print(f"  Negative : {len(df_neg)}")
    print(f"\n✅ Saved → {XLSX_FILE}")

# ── MAIN ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    raw, cutoff_dt, today = fetch_reviews()
    if not raw:
        print("No reviews fetched in the last 7 days.")
        with open("NO_REVIEWS_FLAG.txt", "w") as f:
            f.write("no reviews")
        exit(0)
    rows = parse_reviews(raw)
    save_to_xlsx(rows, cutoff_dt, today)
    print("\n" + "=" * 60)
    print("Pipeline complete! iOS_WearablesApp_Review.xlsx ready.")
    print("=" * 60)
