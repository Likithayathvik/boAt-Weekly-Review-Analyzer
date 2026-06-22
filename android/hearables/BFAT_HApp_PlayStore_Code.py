#Project Name : boAt - App Feedback Analyzer Tool
#Author       : boAt R&D (QA)
#Description  : Google Play Developer API based boAt App Feedback Analyzer
#Note         : Legacy Selenium web-scraping functions have been removed.
#               Review fetching is now handled in main.py via the official
#               Google Play Developer API. Only post-processing, keyword
#               classification, and Excel formatting functions remain below.
#********************************************************************#

import os
import re
import unicodedata
import openpyxl
import pandas as pd
from transformers import pipeline
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment

DEVICE_TYPES = ["Stone", "Rockerz", "Airdopes Atom", "Airdopes Prime", "Airdopes", "Nirvana", "Party Pal", "immortal", "avante bar", "aavante", "party pal", "unity"]

# Alias map: normalise fuzzy-matched prefix → canonical display name
DEVICE_TYPE_ALIAS = {
    "aavante": "Aavante Bar",
    "avante bar": "Aavante Bar",
    "unity": "Boat Unity",
    "airpods":"Airdopes",
}

CATEGORY_KEYWORDS = ["battery information","charging information","Crash","battery", "Close","connect", "custom equalizer","Equalizer","Diagnostics",
                     "Login","log into","log", "Location", "tap", "button", "touch","EQ", "fw update","firmware update","Update",
                     "phone number","mobile number","otp",
                     "noise cancellation","anc",
                     "freez","privacy","policy","account","screen","download","load","stuck","support","warranty","app open","warranty","service","details",
                     "mydetails","personal information","personal","email address","email",
                     "control",
                     "compatible","customer care","complaint","ldac","tablet","tab","ipad",]

PRIORITY_KEYWORDS = [
    "custom equalizer",
    "tablet",
    "tab",
    "ipad",
    "log in",
    "firmware update",
    "fw update",
    "battery information",
    "charging information",
    "login",
    "app open",
    "phone number",
    "mobile number",
    "otp",
    "personal information",
    "email address",
    "mydetails",
    "connect",
    "connectivity",
    "sku not supported",
]

FORCE_UPDATE_SIGNALS = [
    "without updating", "without update", "let me use",
    "skip update", "force update", "forced to update",
    "dont want to update", "don't want to update",
    "not updating", "without upgrading"
]

THRESHOLD = 80
# Global variables declaration
global App_Rating
global App_Version
global lowest_date
global latest_date

products = [
    "Rockerz 480","Rockerz 421","Airdopes Ace Gen 2","Airdopes Alpha Gen 2",
    "Airdopes 213","Airdopes 212","Airdopes 219","Rockerz 113","Rockerz 112",
    "Rockerz 111","Airdopes Prime 513 ANC","Rockerz Zen ANC","Airdopes 131 Gen 2",
    "Rockerz Prime 415","Rockerz 301 ANC","Rockerz Plus 550 ANC","Rockerz 512 ANC",
    "Stone Arc Pro Plus","Stone Arc Pro","Aavante Bar 950","Aavante Bar Aspire",
    "Aavante Bar 490","Aavante Bar 485","Aavante Bar 480","Airdopes Prime 413",
    "Rockerz Trinity Grande","Airdopes 313","Rockerz Prime 255Z","Rockerz Prime 205",
    "Airdopes Prime 511","Airdopes Prime 412","Stone 358","Stone 352","Stone 350",
    "Stone 358 Pro","Stone 352 Pro","Stone 350 Pro","Airdopes 301","Airdopes 181 Pro",
    "Airdopes 161 ANC Elite","Rockerz Summit","Airdopes 91 Prime","Airdopes 141 Elite ANC",
    "Airdopes Ace","Airdopes 71","Airdopes Ultra Plus","Airdopes 131 Elite ANC",
    "Airdopes 101 v2","Airdopes Joy","Rockerz 235 Pro","Rockerz 430","Airdopes 121 Pro Plus",
    "Airdopes Plus 311","Airdopes 311 Pro","Airdopes 111v2","Rockerz Trinity Gen 2",
    "Airdopes Prime 701 ANC","Rockerz 450 Pro","Rockerz 450R","Rockerz 450",
    "Nirvana Ivy Pro","Nirvana Zenith Pro","Nirvana Iris","PartyPal 600","Rockerz 650 Pro",
    "Rockerz 551 ANC Pro","Nirvana X TWS","Airdopes 800 HiDef","Nirvana Crystl",
    "Airdopes Ultra Pro","Nirvana Lucid","Airdopes Alpha","Airdopes 161","Airdopes 141",
    "Airdopes 131","Nirvana Ion ANC Pro","Rockerz 210 ANC","Nirvana Ivy","Rockerz 255 Z Plus",
    "Stone Lumos","Nirvana Space","NIRVANA ZENITH","NIRVANA NEBULA","Nirvana Eutopia",
    "Airdopes Supreme","Rockerz 255 ANC","Nirvana Ion ANC","Nirvana Ion","Airdopes Flex 454 ANC",
    "Airdopes 800","Rockerz 255 Max","Airdopes 300","NIRVANA 525ANC","Airdopes 341ANC",
    "Airdopes 393ANC","Airdopes 172","Rockerz 255 Pro+","Rockerz 333 Pro","Rockerz 333",
    "Rockerz 330 Pro","Nirvana Crown","Airdopes Prime 700 ANC",
]

# Map Reason keywords to Consolidated Reason
CONSOLIDATION_MAP = {
    "custom equalizer": "Custom Equalizer issue",
    "equalizer": "Equalizer issue",
    "eq": "Equalizer issue",
    "option": "Button Customization",
    "tap": "Button Customization",
    "touch": "Button Customization",
    "button": "Button Customization",
    "customization": "Button Customization",
    "control": "Button Customization",
    "crash": "App Crash",
    "close": "App Crash",
    "app close": "App Crash",
    "connect": "Connectivity",
    "detect": "Connectivity",
    "fw update": "Firmware Update",
    "firmware update": "Firmware Update",
    "battery information": "Battery Information",
    "charging information": "Battery Information",
    "battery": "Battery Information",
    "noise cancellation": "Noise Cancellation Issue",
    "anc": "Noise Cancellation Issue",
    "login": "Login issue",
    "log into": "Login issue",
    "log": "Login issue",
    "issue": "App Issue",
    "app open": "App Issue",
    "phone number": "App Privacy Issue",
    "mobile number": "App Privacy Issue",
    "otp": "App Privacy Issue",
    "personal information": "App Privacy Issue",
    "email address": "App Privacy Issue",
    "email": "App Privacy Issue",
    "location": "Location Permission",
    "freez": "App Crash",
    "positive": "Positive Feedback",
    "update": "App Update issue",
    "app space": "App size",
    "download": "App download Issue",
    "privacy": "App Privacy Issue",
    "policy": "App Policy Issue",
    "account": "Login Issue",
    "screen": "App Issue",
    "load": "App Issue",
    "stuck": "App Issue",
    "support": "SKU not supported",
    "diagnostics": "Smart diagnostics",
    "warranty": "Service Issue",
    "service": "Service Issue",
    "details": "Privacy Issue",
    "mydetails": "App Privacy Issue",
    "personal": "App Privacy Issue",
    "compatible": "App not supported",
    "customer care": "Service Issue",
    "complaint": "Service Issue",
    "ldac": "App feature Issue",
    "tablet": "App not supported",
    "tab": "App not supported",
    "ipad": "App not supported",
}

from rapidfuzz import fuzz, process
import re

# Global variables declaration
global App_Rating
global App_Version
global lowest_date
global latest_date

#***********************************************************************#

def Positive_Reviews():
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_HearablesApp_Review.xlsx')

    # Filter the DataFrame to keep only rows with ratings 4 and 5
    positive_reviews_df = df[df['Ratings'].isin([4, 5])]

    # Save the positive reviews to the same Excel file in a new sheet named 'Positive_Reviews'
    with pd.ExcelWriter('Android_HearablesApp_Review.xlsx', engine='openpyxl', mode='a') as writer:
        positive_reviews_df.to_excel(writer, sheet_name='Positive_Reviews', index=False)

    print("Positive reviews saved to 'Android_HearablesApp_Review.xlsx' under sheet name 'Positive_Reviews'.")


def normalize_unicode_text(text):
    """Convert stylized Unicode characters (𝓷𝓲𝓬𝓮 → nice) to plain ASCII."""
    return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')



def extract_device_names(text, threshold=70):
    """
    Device extraction with:
    ✔ Fuzzy match only for prefix (Device Type)
    ✔ Keep numeric model EXACT — picks first number AFTER prefix and BEFORE suffix
    ✔ EXACT suffix detection
    ✔ Prefix + alpha words (no number): e.g. "Stone Spinx Pro", "Nirvana Lucid"
    ✔ Aavante-specific: captures Prefix + version (e.g. 2.0) + model number
      e.g. "Boat Aavante 2.0 150" → "Aavante Bar 2.0 150"
      Handles "aavante" (double-a) spelling via DEVICE_TYPE_ALIAS.
    """
    if not text or not isinstance(text, str):
        return "Unknown Device"
    if len(text.split()) <= 2:
        return "Unknown Device"

    txt = text.lower()

    # ── Capture version strings (e.g. "2.0") BEFORE tokenising ──────────────
    version_strings = re.findall(r'\b\d+\.\d+\b', txt)   # e.g. ["2.0"]

    # 1️⃣ Fuzzy match PREFIX only
    words = re.findall(r"[a-zA-Z0-9]+", txt)
    windows = words + [" ".join(words[i:i+2]) for i in range(len(words)-1)]
    best_prefix = None
    best_score = 0
    for w in windows:
        result = process.extractOne(
            w, [p.lower() for p in DEVICE_TYPES], scorer=fuzz.ratio
        )
        if result and result[1] >= threshold and (
            result[1] > best_score or
            (result[1] == best_score and best_prefix and len(result[0]) > len(best_prefix))
        ):
            best_prefix = result[0]
            best_score = result[1]

    # Resolve alias → canonical display name (e.g. "aavante" → "Aavante Bar")
    canonical_prefix = DEVICE_TYPE_ALIAS.get(best_prefix, best_prefix.title() if best_prefix else None)

    # 2️⃣ Strip version strings, then find all numbers with their positions
    txt_no_version = re.sub(r'\b\d+\.\d+\b', '', txt)
    nums = re.findall(r"\d+", txt_no_version)

    SUFFIX_MAP = {
        "anc elite": "ANC Elite",
        "anc pro": "ANC Pro",
        "ancpro": "ANC Pro",
        "anc": "ANC",
        "pro+": "Pro+",
        "pro plus": "Pro+",
        "pro": "Pro",
        "plus": "Plus",
        "prime": "Prime",
        "gen 2": "Gen 2",
        "gen2": "Gen 2",
        "elite": "Elite",
        "ultra pro": "Ultra Pro",
        "ultra": "Ultra",
        "max": "Max",
        "v1": "v1",
        "v2": "v2",
        "v2 pro": "V2 Pro",
        "v2pro": "V2 Pro",
        "r": "R",
        "z plus": "Z Plus",
        "zplus": "Z Plus",
        "z": "Z",
        "TWS": "TWS",
        "x tws": "X TWS",
        "xtws": "X TWS",
        "hidef": "HiDef",
        "hi def": "HiDef",
        "ionanc": "Ion ANC",
        "ion anc": "Ion ANC",
        "ion anc pro": "Ion Anc Pro",
        "ionancpro": "Ion Anc Pro",
        "ion": "ion",
        "loop": "loop",
        "tws": "TWS",
        "v2Pro": "V2 Pro",
        "v2 Pro": "V2 Pro",
        "crown": "Crown",
        "ivy": "Ivy",
        "IvyPro": "IvyPro",
        "Ivy Pro": "Ivy Pro",
    }

    # ── Pick model_num: first number that appears AFTER the prefix
    #    AND BEFORE the earliest suffix match in the text ──────────────────────
    model_num = None
    if nums and best_prefix:
        # Find where the prefix ends in txt_no_version
        prefix_first_word = best_prefix.split()[0]
        prefix_match = re.search(re.escape(prefix_first_word), txt_no_version)
        prefix_end_pos = prefix_match.end() if prefix_match else 0

        # Find where the earliest suffix begins in txt_no_version
        suffix_start_pos = len(txt_no_version)  # default: end of string
        for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
            m = re.search(r"\b" + re.escape(key) + r"\b", txt_no_version)
            if m and m.start() > prefix_end_pos:
                if m.start() < suffix_start_pos:
                    suffix_start_pos = m.start()

        # All numbers with positions
        num_positions = [(m.group(), m.start()) for m in re.finditer(r"\d+", txt_no_version)]

        # Pick the first number that falls strictly between prefix_end and suffix_start
        between = [(n, pos) for n, pos in num_positions
                   if pos >= prefix_end_pos and pos < suffix_start_pos]
        if between:
            model_num = between[0][0]
        else:
            # Fallback: first number after prefix (ignore suffix boundary)
            if suffix_start_pos == len(txt_no_version):
                after_prefix = [(n, pos) for n, pos in num_positions if pos >= prefix_end_pos]
                if after_prefix:
                    model_num = after_prefix[0][0]
            # No number after prefix at all → model_num stays None
    elif nums:
        # No prefix found — keep old largest-number fallback
        model_num = sorted(nums, key=lambda x: (len(x), int(x)), reverse=True)[0]

    # ── Aavante-specific: Prefix + version + model number ────────────────────
    if best_prefix and "avante" in best_prefix.lower() and version_strings and model_num:
        version_str = version_strings[0]   # e.g. "2.0"
        detected_suffix = ""
        for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
            if re.search(r"\b" + re.escape(key) + r"\b", txt):
                detected_suffix = SUFFIX_MAP[key]
                break
        device = f"{canonical_prefix} {version_str} {model_num}"
        if detected_suffix:
            device += f" {detected_suffix}"
        return device.strip()

    # -------------------------
    #  Prefix + alpha + optional suffix (no number in text)
    # -------------------------
    if best_prefix and not model_num:
        prefix_tokens = best_prefix.split()
        prefix_end = -1

        # Locate where the prefix ends in the word list
        for i in range(len(words)):
            window = " ".join(words[i:i + len(prefix_tokens)])
            if fuzz.ratio(window, best_prefix) >= threshold:
                prefix_end = i + len(prefix_tokens)
                break

        if prefix_end != -1:
            # --- detect the longest matching suffix present in text ---
            detected_suffix = ""
            for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
                if re.search(r"\b" + re.escape(key) + r"\b", txt):
                    detected_suffix = SUFFIX_MAP[key]
                    break

            suffix_tokens = detected_suffix.lower().split() if detected_suffix else []
            middle_parts = []
            for w in words[prefix_end:]:
                # stop if we've reached the suffix word(s)
                if suffix_tokens and w == suffix_tokens[0]:
                    break
                # collect only alphabetic words (skip noise / numbers)
                if re.match(r'^[a-zA-Z]{2,}$', w):
                    middle_parts.append(w)
                    # no suffix expected → one model-name word is enough, stop
                    if not detected_suffix:
                        break
                else:
                    break   # stop at any digit/punctuation token

            parts = [canonical_prefix]
            if middle_parts:
                parts.append(" ".join(w.title() for w in middle_parts))
            if detected_suffix:
                parts.append(detected_suffix)

            result_device = " ".join(parts).strip()
            return result_device

        # Alpha-only fallback: fuzzy match full review against products list
        fallback = process.extractOne(
            txt, [p.lower() for p in products], scorer=fuzz.partial_ratio
        )
        if fallback and fallback[1] >= THRESHOLD:
            matched_idx = [p.lower() for p in products].index(fallback[0])
            return products[matched_idx]
        return "Unknown Device"

    if not best_prefix:
        detected_suffix = ""
        for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
            for num in nums:
                pattern = rf"{num}\s*{re.escape(key)}\b"
                if re.search(pattern, txt, flags=re.IGNORECASE):
                    detected_suffix = SUFFIX_MAP[key]
                    break
            if detected_suffix:
                break
        if not detected_suffix:
            for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
                for num in nums:
                    pattern = rf"{num}{re.escape(key)}\b"
                    if re.search(pattern, txt, flags=re.IGNORECASE):
                        detected_suffix = SUFFIX_MAP[key]
                        break
                if detected_suffix:
                    break
        if detected_suffix:
            return f"{model_num} {detected_suffix}".strip()
        # No-prefix fallback: fuzzy match full review against products list
        fallback = process.extractOne(
            txt, [p.lower() for p in products], scorer=fuzz.partial_ratio
        )
        if fallback and fallback[1] >= THRESHOLD:
            matched_idx = [p.lower() for p in products].index(fallback[0])
            return products[matched_idx]
        return "Unknown Device"

    detected_suffix = ""
    for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
        if re.search(r"\b" + re.escape(key) + r"\b", txt):
            detected_suffix = SUFFIX_MAP[key]
            break
    if not detected_suffix:
        for num in nums:
            for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
                pattern = rf"{num}{re.escape(key)}\b"
                if re.search(pattern, txt):
                    detected_suffix = SUFFIX_MAP[key]
                    break
            if detected_suffix:
                break

    # 4️⃣ Build device
    device = f"{canonical_prefix} {model_num}" if canonical_prefix else "Unknown Device"
    if detected_suffix:
        device += f" {detected_suffix}"

    # 5️⃣ FINAL FALLBACK: if still Unknown Device, fuzzy-match the full review
    #    text directly against the products list.
    if device.strip() == "Unknown Device":
        result = process.extractOne(
            txt, [p.lower() for p in products], scorer=fuzz.partial_ratio
        )
        if result and result[1] >= THRESHOLD:
            matched_idx = [p.lower() for p in products].index(result[0])
            device = products[matched_idx]

    return device.strip()


# Remove emojis before sentiment analysis

def remove_emojis(text):
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"  # emoticons
        "\U0001F300-\U0001F5FF"  # symbols & pictographs
        "\U0001F680-\U0001F6FF"  # transport & map symbols
        "\U0001F1E0-\U0001F1FF"  # flags (iOS)
        "\U00002700-\U000027BF"  # dingbats
        "\U0001F900-\U0001F9FF"  # supplemental symbols
        "\U00002600-\U000026FF"  # miscellaneous symbols
        "]+",
        flags=re.UNICODE,
    )
    return emoji_pattern.sub("", text)


EXCEPTION_PHRASES = ["even though", "even", "though","issue"]

def contains_exception_phrase(text):
    """
    Returns True if any exception phrase matches fuzzily above threshold.
    """
    text_lower = text.lower()
    for phrase in EXCEPTION_PHRASES:
        # Sliding window over words to check fuzzy match
        words = text_lower.split()
        for i in range(len(words)):
            # check up to len(phrase.split())+1 words together
            window = " ".join(words[i:i+len(phrase.split())+1])
            score = fuzz.partial_ratio(phrase, window)
            if score >= THRESHOLD:
                return True
    return False


def is_valid_device(d, products_set):
    """
    Check if device d is valid.
    - If numeric: match only against numeric parts of products_set
    - Otherwise: exact match in products_set
    """
    if d.isdigit():
        for product in products_set:
            # Extract all numeric parts from product name
            nums = re.findall(r'\d+', product)
            if d in nums:
                return True
        return False
    else:
        return d in products_set


def is_sentiment_positive(review, sentiment_analyzer):
    if not isinstance(review, str):
        return False
    cleaned = remove_emojis(review).strip()
    if not cleaned or len(cleaned.split()) > 5:
        return False
    try:
        return sentiment_analyzer(cleaned)[0]['label'].lower() == 'positive'
    except Exception:
        return False



def best_keyword_match(review_lower, keywords, threshold):
    """
    Three-tier matching:
      Tier 1 (priority exact): keyword is in PRIORITY_KEYWORDS AND is a literal
                               substring of the review. First match in
                               PRIORITY_KEYWORDS order wins immediately.
                               Also fuzzy-matches multi-word priority keywords
                               for typo resilience.
      Tier 2 (exact substring): keyword appears literally but is NOT a priority kw.
                               Among all tier-2 hits, longest keyword wins.
      Tier 3 (fuzzy only):     no exact match found at all; highest partial_ratio
                               wins (ties broken by keyword length).
    Returns "Other" if no keyword reaches `threshold`.
    """
    # Tier 1: priority keywords — exact first, then fuzzy for multi-word (typo resilience)
    for pkw in PRIORITY_KEYWORDS:
        pkw_lower = pkw.lower()
        if pkw_lower in review_lower:
            return pkw_lower
        if len(pkw_lower.split()) > 1:
            score = fuzz.partial_ratio(pkw_lower, review_lower)
            if score >= threshold:
                return pkw_lower

    # Tier 2: exact substring, non-priority — longest wins
    exact_kw  = None
    exact_len = 0

    # Tier 3: fuzzy fallback
    fuzzy_kw    = "Other"
    fuzzy_score = 0
    fuzzy_len   = 0

    for kw in keywords:
        kw_lower = kw.lower()
        if kw_lower in review_lower:
            if len(kw_lower) > exact_len:
                exact_kw  = kw_lower
                exact_len = len(kw_lower)
            continue

        score = fuzz.partial_ratio(kw_lower, review_lower)
        if score < threshold:
            continue
        if score > fuzzy_score or (score == fuzzy_score and len(kw_lower) > fuzzy_len):
            fuzzy_kw    = kw_lower
            fuzzy_score = score
            fuzzy_len   = len(kw_lower)

    return exact_kw if exact_kw is not None else fuzzy_kw



def Negative_Reviews():
    df = pd.read_excel('Android_HearablesApp_Review.xlsx')

    # ── Normalize stylized Unicode in Reviews column so readable text is stored ──
    df['Reviews'] = df['Reviews'].apply(
        lambda x: normalize_unicode_text(str(x)).strip() if pd.notna(x) else x
    )

    negative_reviews_df = df[df['Ratings'].isin([1,2,3])].copy()
    sentiment_analyzer = pipeline("sentiment-analysis")

    reasons, DUT = [], []
    for review in negative_reviews_df['Reviews']:
        # ── Empty / NaN ──────────────────────────────────────────────────────
        if pd.isna(review) or not str(review).strip():
            reasons.append("Generic Comments")
            DUT.append("Unknown Device")
            continue

        if is_sentiment_positive(review, sentiment_analyzer):
            reasons.append("positive")
            DUT.append(extract_device_names(review))
            continue

        review_lower = review.lower()
        matched_category = best_keyword_match(review_lower, CATEGORY_KEYWORDS, THRESHOLD)

        # ── Fallback to full sentiment if no keyword matched ──────────────────
        if matched_category == "Other":
            clean_review = remove_emojis(review)
            if contains_exception_phrase(clean_review):
                sentiment = "neutral"
            else:
                sentiment = sentiment_analyzer(clean_review)[0]['label'].lower()
            if sentiment == "positive":
                matched_category = "positive"

        reasons.append(matched_category)
        DUT.append(extract_device_names(review))

    negative_reviews_df['Reason'] = reasons
    negative_reviews_df['DUT'] = DUT

    # ── Consolidation ─────────────────────────────────────────────────────────
    consolidated_reasons = []
    products_set = set(products)
    reviews_list = negative_reviews_df['Reviews'].tolist()

    for review, reason, dut in zip(reviews_list, reasons, DUT):
        r = reason.lower()
        d = dut.strip()
        review_lower = str(review).lower() if pd.notna(review) else ""

        # ── SKU not supported: device extracted but not in known products list ──
        if d != "Unknown Device" and not is_valid_device(d, products_set):
            consolidated = "SKU not supported"

        # ── Force update vs regular update issue ──────────────────────────────
        elif r == "update":
            if any(sig in review_lower for sig in FORCE_UPDATE_SIGNALS):
                consolidated = "Force App Update Not Required"
            else:
                consolidated = "App Update issue"

        # ── Standard consolidation map ─────────────────────────────────────────
        else:
            consolidated = CONSOLIDATION_MAP.get(r, "Generic Comments")

        consolidated_reasons.append(consolidated)

    negative_reviews_df['Consolidated Reason'] = consolidated_reasons

    with pd.ExcelWriter('Android_HearablesApp_Review.xlsx', engine='openpyxl', mode='a') as writer:
        negative_reviews_df.to_excel(writer, sheet_name='Negative_Reviews', index=False)

    print("Negative reviews and reasons saved to 'Negative_Reviews' sheet.")


def SheetNameChange():
    directory_path = os.path.dirname(os.path.realpath(__file__))  # Get the current script's directory
    # Load the Excel file
    file_path = 'Android_HearablesApp_Review.xlsx'
    current_sheet_name = 'Sheet1'
    new_sheet_name = 'Happ_Review_Data'

    try:
        # Load the Excel file
        wb = openpyxl.load_workbook(file_path)

        # Check if the current sheet name exists
        if current_sheet_name in wb.sheetnames:
            # Get the sheet object
            sheet = wb[current_sheet_name]

            # Change the sheet name to 'Negative_Reviews'
            sheet.title = new_sheet_name

            # Save the modified workbook with the new sheet name
            wb.save(file_path)

            print(f"The sheet name '{current_sheet_name}' has been changed to '{new_sheet_name}'.")
        else:
            print(f"The sheet '{current_sheet_name}' does not exist in the Excel file.")

        # Close the Excel file
        wb.close()

    except FileNotFoundError:
        print(f"The file '{file_path}' does not exist.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")



def Sorting_Xlsx(sheet_name):
    import math

    # Load the XLSX file
    workbook = openpyxl.load_workbook('Android_HearablesApp_Review.xlsx')

    # Select the worksheet where you want to set column widths
    worksheet = workbook[sheet_name]

    # ── FIX: column widths are now keyed by HEADER NAME, not hardcoded
    #         letters. The old dict only covered A-L, so any sheet with
    #         more columns (e.g. Negative_Reviews, which has Reason, DUT,
    #         Consolidated Reason, AI_Confidence, AI_Consolidated_Reason,
    #         QA Comments after boAt Reply) fell back to Excel's default
    #         ~8.43 width for everything past column L, and column L's
    #         width (tuned for one sheet) got reused for a totally
    #         different column on the other sheet. That's why
    #         Negative_Reviews looked clumpsy compared to Happ_Review_Data.
    default_width = 18
    width_by_header = {
        'reviewid': 15,
        'user name': 15,
        'ratings': 8,
        'reviews': 42,
        'phone name': 12.5,
        'os version': 11,
        'date time': 11,
        'version code': 10,
        'app version': 10,
        'boat reply': 48,
        'reason': 15,
        'dut': 21,
        'consolidated reason': 30,
        'ai_confidence [testing purpose]': 12,
        'ai_consolidated_reason [testing purpose]': 30,
        'qa comments [give the comments based on consolidated reason]': 40,
    }

    # Build column_widths dynamically from the sheet's actual header row,
    # so EVERY column present gets a sensible width — not just A-L.
    # FIX: some headers (e.g. "AI_Confidence\n[Testing Purpose]") contain
    # embedded newlines for display wrapping in the header cell itself —
    # normalize all whitespace/newlines to single spaces before lookup,
    # or the dict match silently fails and falls back to default_width.
    header_cells = next(worksheet.iter_rows(min_row=1, max_row=1))
    column_widths = {}
    for cell in header_cells:
        if cell.value is None:
            continue
        header_key = " ".join(str(cell.value).split()).strip().lower()
        column_widths[cell.column_letter] = width_by_header.get(header_key, default_width)

    # Set the column widths
    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width

    # Iterate through the rows and columns to apply wrap text and middle align
    for row in worksheet.iter_rows():
        for cell in row:
            # Set wrap text to True
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")

    # ── FIX: wrap_text alone doesn't expand row height in Excel — we need
    #         to explicitly calculate and set row heights based on the
    #         longest wrapped text in each row, otherwise wrapped cells
    #         look clipped/clumsy even though wrap is technically on.
    def estimate_row_height(text, col_width, font_size=11):
        if not isinstance(text, str) or not text.strip():
            return 15
        # Rough chars-per-line estimate based on column width
        chars_per_line = max(1, int(col_width * 1.8))
        lines = 0
        for line in text.split("\n"):
            lines += max(1, math.ceil(len(line) / chars_per_line))
        return min(409, max(15, lines * (font_size + 4)))

    # FIX: iterate over ALL columns that have a width set (i.e. every real
    # column in this sheet), not just the old hardcoded A-L subset — so
    # row height correctly accounts for long text in M, N, O, P, etc.
    for row_idx in range(2, worksheet.max_row + 1):
        max_height = 15
        for col_letter, width in column_widths.items():
            cell = worksheet[f"{col_letter}{row_idx}"]
            if cell.value:
                height = estimate_row_height(str(cell.value), width)
                max_height = max(max_height, height)
        worksheet.row_dimensions[row_idx].height = max_height

    # Header row gets a fixed comfortable height
    worksheet.row_dimensions[1].height = 30

    # Save the workbook with changes to the same file
    workbook.save('Android_HearablesApp_Review.xlsx')

    # Close the workbook
    workbook.close()

    print(f"Xlsx modification done for sheet '{sheet_name}'.")


def Review_Date_Calculator():
    global lowest_date
    global latest_date
    # Load the Excel file
    df = pd.read_excel('Android_HearablesApp_Review.xlsx')

    # Replace 'DateTimeColumn' with the actual name of the column containing your date and time data
    df['Date Time'] = pd.to_datetime(df['Date Time'])

    # Extract date part
    df['DateOnly'] = df['Date Time'].dt.date

    # Find the earliest and latest date
    lowest_date = df['DateOnly'].min()
    latest_date = df['DateOnly'].max()


    # Print the results
    print("Lowest Date in the DateTimeColumn:", lowest_date)
    print("Latest Date in the DateTimeColumn:", latest_date)



def Negative_Review_Analysis():
    # Load the Excel file into a pandas DataFrame
    df = pd.read_excel('Android_HearablesApp_Review.xlsx', sheet_name='Negative_Reviews')

    # Count occurrences of each "Consolidated Reason"
    reason_counts = df['Consolidated Reason'].value_counts()

    # Print the counts for each "Consolidated Reason"
    print("Negative Review Analysis:")
    for reason, count in reason_counts.items():
        print(f"{reason}: {count}")

    return reason_counts


def Summary():
    global App_Rating
    global App_Version
    global lowest_date
    global latest_date
    import pandas as pd
    from openpyxl import load_workbook
    from openpyxl.utils.dataframe import dataframe_to_rows
    from openpyxl.styles import Alignment, Border, Side, Font, PatternFill
    import math

    print("Generating styled Summary1 sheet...")

    # Load main review data
    df = pd.read_excel('Android_HearablesApp_Review.xlsx')

    # Review counts
    pos_count = len(df[df['Ratings'].isin([4, 5])])
    neg_count = len(df[df['Ratings'].isin([1, 2, 3])])

     # Calculate summary statistics
    app_name = "Android Hearables App " + str(App_Version)	# Replace with your app name
    app_rating = App_Rating + '*'
    review_data = str(lowest_date) + ' to ' + str(latest_date)

    # Define app details data
    app_details_data = {
        'App Name': [app_name],
		'Latest App Version': [App_Version],  # Keep the app version unchanged
        'Review Date': [review_data]
    }
    app_df = pd.DataFrame(app_details_data)
    counts_df = pd.DataFrame({
        'Category': ['Positive Reviews (4 & 5)', 'Negative Reviews (<=3)'],
        'Count': [pos_count, neg_count]
    })

    # Load Negative Reviews
    neg_df = pd.read_excel('Android_HearablesApp_Review.xlsx', sheet_name='Negative_Reviews')
    reason_counts = neg_df['Consolidated Reason'].value_counts()
    summary_df = pd.DataFrame({
        'Summary (-ve reviews)': reason_counts.index,
        'Count': reason_counts.values
    })

    # Helper: multiline list excluding Unknown Device
    def make_multiline(df_filter):
        df_filter = df_filter[df_filter['DUT'] != 'Unknown Device']
        if df_filter.empty:
            return "No valid devices found"
        counts = df_filter['DUT'].value_counts()
        return "\n".join([f"{dev} - {cnt}" for dev, cnt in counts.items()])

    # Remarks column
    summary_df['Remarks'] = [
        make_multiline(neg_df[neg_df['Consolidated Reason'] == reason])
        for reason in summary_df['Summary (-ve reviews)']
    ]

    if 'Positive Feedback' in summary_df['Summary (-ve reviews)'].values:
        summary_df.loc[
            summary_df['Summary (-ve reviews)'] == 'Positive Feedback', 'Remarks'
        ] = "Reviews are good but ratings are <= 3"

    # Workbook setup
    wb = load_workbook('Android_HearablesApp_Review.xlsx')
    ws = wb['Summary'] if 'Summary' in wb.sheetnames else wb.create_sheet('Summary')
    ws.delete_rows(1, ws.max_row)

    # Title
    ws.append(["📊 Review Summary Report"])
    ws.merge_cells('A1:C1')
    title = ws['A1']
    title.font = Font(size=14, bold=True, color="FFFFFF")
    title.alignment = Alignment(horizontal='center', vertical='center')
    title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")

    ws.append([])

    # Write all data sections
    for df_part in [app_df, pd.DataFrame(), counts_df, pd.DataFrame(), summary_df]:
        for r in dataframe_to_rows(df_part, index=False):
            ws.append(r)

    # Basic styles
    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical='top')
    blue_header = PatternFill(start_color="A7C7E7", end_color="A7C7E7", fill_type="solid")
    green_section = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")

    # 📏 Increased Column C width
    col_widths = {'A': 25, 'B': 15, 'C': 35, 'D': 23}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    # Apply border + wrap
    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = wrap

    # Highlight headers/sections
    for row in ws.iter_rows(min_row=1, max_col=3):
        values = [str(c.value).strip().lower() if c.value else "" for c in row]
        if "category" in values or "summary (-ve reviews)" in values:
            for cell in row:
                cell.fill = blue_header
                cell.font = Font(bold=True)
        elif "app name" in values:
            for cell in row:
                cell.fill = green_section
                cell.font = Font(bold=True)

    # 🧭 Center align Column C (Remarks)
    for row in ws.iter_rows():
        for cell in row:
            if hasattr(cell, "column_letter") and cell.column_letter == 'C':
                cell.alignment = Alignment(
                    wrap_text=True, horizontal='center', vertical='center'
                )

    # Dynamic row height for Column C
    def calc_row_height(text, col_width):
        if not isinstance(text, str) or not text.strip():
            return 15
        lines = text.split("\n")
        est_lines = sum(max(1, math.ceil(len(line) / (col_width * 1.5))) for line in lines)
        return min(300, 15 * est_lines)

    col_c_width = ws.column_dimensions['C'].width or 35
    for row in ws.iter_rows(min_row=1, max_col=3):
        for cell in row:
            if not hasattr(cell, "column_letter"):
                continue
            if cell.column_letter == 'C':
                ws.row_dimensions[cell.row].height = calc_row_height(cell.value, col_c_width)

    wb.save('Android_HearablesApp_Review.xlsx')
    print("✅ Summary sheet updated — Column C centered, wider, and formatted cleanly.")


def close_excel_file(file_path):
    """
    Close the Excel file to release any locks on it.
    """
    try:
        workbook = openpyxl.load_workbook(file_path)
        workbook.close()
        print("Closed the Excel file.")
    except Exception as e:
        print(f"Error closing the Excel file: {e}")

