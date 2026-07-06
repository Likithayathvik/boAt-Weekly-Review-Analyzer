#Project Name : boAt - App Feedback Analyzer Tool
#Author       : boAt R&D (QA)
#Date         : 15 April 2024
#Version      : v2.3 (Updated)
#Description  : Google Play Developer API based boAt App Feedback Analyzer
#Updated      : Consolidated Equalizer and ANC for Airhoa SDK 16KB page support
#Note         : Legacy Selenium web-scraping functions have been removed.
#               Review fetching is now handled in main.py via the official
#               Google Play Developer API. Only post-processing, keyword
#               classification, and Excel formatting functions remain below.
#********************************************************************#

import os
import re
import math
import unicodedata
import openpyxl
import pandas as pd
from transformers import pipeline
from rapidfuzz import fuzz, process
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

PREFIXES = [
    "Airdopes Plus Spark",
    "Airdopes Plus",
    "Rockerz Zen",
    "Rockerz Prime",
    "Rockerz Plus",
    "Airdopes Prime",
    "Stone Arc",
    "Stone",
    "PartyPal",
    "Nirvana Ion",
    "Nirvana",
    "Rockerz Trinity",
    "Airdopes Ultra",
    "Airdopes Flex",
    "Rockerz",
    "Immortal",
    "Airdopes Ace",
    "Airdopes Alpha",
    "Airdopes",
    "Aavante Bar",
    "Stone Spinx",
    "Airdopes Bliss",
    "Airdopes Unity",
]

DEVICE_TYPE_ALIAS = {
    "aavante": "Aavante Bar",
    "avante bar": "Aavante Bar",
    "unity": "Boat Unity",
    "airpods": "Airdopes",
}

CATEGORY_KEYWORDS = ["battery information","charging information","Crash","battery", "Close","connect", "custom equalizer","Equalizer","Diagnostics",
                     "Login","log into","log", "Location", "tap", "button", "touch","EQ", "fw update","firmware update","Update",
                     "phone number","mobile number","otp",
                     "noise cancellation","anc",
                     "freez","privacy","policy","account","screen","download","install","load","stuck","support","warranty","app open","warranty","service","details",
                     "mydetails","personal information","personal","email address","email",
                     "control",
                     "compatible","customer care","complaint","service","ldac","tablet","tab","ipad","internet", "offline",]

PRIORITY_KEYWORDS = [
    "custom equalizer",
    "anc",
    "equalizer",
    "eq",
    "noise cancellation",
    "tablet",
    "tab",
    "ipad",
    "button",
    "battery information",
    "charging information",
    "battery",
    "percentage",
    "mobile number",
    "log in",
    "login",
    "firmware update",
    "fw update",
    "app open",
    "otp",
    "personal information",
    "email address",
    "mydetails",
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

SDK_PRODUCTS = [
    "Airdopes Plus 421 ANC", "Airdopes Plus 224", "Airdopes Plus Spark ANC", "Airdopes Plus Wave",
    "Airdopes Plus Neo", "Airdopes Plus 222", "Airdopes ProClip", "Airdopes Plus 223",
    "Rockerz Plus 450 ANC", "Airdopes Prime 700 ANC", "Rockerz Zen ANC", "Rockerz Prime 415",
    "Rockerz 301 ANC", "Rockerz Plus 550 ANC", "Airdopes Prime 512", "Rockerz 512 ANC",
    "Stone Arc Pro Plus", "Stone Arc Pro", "Nirvana Crown", "Airdopes Prime 413",
    "Rockerz Trinity Grande", "Rockerz Prime 255Z", "Rockerz Prime 205", "Airdopes Prime 511",
    "Airdopes Prime 412", "Airdopes Prime 513 ANC", "Airdopes Prime 701 ANC", "Nirvana Luxe",
    "Nirvana Ivy Pro", "Nirvana Zenith Pro", "Nirvana Iris", "PartyPal 600", "Rockerz 650 Pro",
    "Rockerz 551 ANC Pro", "Nirvana X TWS", "Airdopes 800 HiDef", "Nirvana Crystl",
    "Airdopes Ultra Pro", "Nirvana Lucid", "Nirvana Ion ANC Pro", "Rockerz 210 ANC",
    "Nirvana Ivy", "Rockerz 255 Z Plus", "Stone Lumos", "Nirvana Space", "NIRVANA ZENITH",
    "NIRVANA NEBULA", "Nirvana Eutopia", "Airdopes Supreme", "Rockerz 255 ANC",
    "Nirvana Ion ANC", "Nirvana Ion", "Airdopes Flex 454 ANC", "Airdopes 800",
    "Rockerz 255 Max", "Airdopes 300", "NIRVANA 525ANC", "Airdopes 341ANC", "Airdopes 393ANC",
    "Airdopes 172", "Rockerz 255 Pro+", "Rockerz Trinity Gen 2", "Rockerz 333 Pro",
    "Rockerz 333", "Rockerz 330 Pro",
]

NON_SDK_PRODUCTS = [
    "Rockerz plus 440", "Airdopes 71 Gen 2", "Airdopes 148 Gen 2", "Airdopes 138",
    "Immortal 131", "Airdopes 163", "Rockerz 150 Pro", "Airdopes 91", "Airdopes 141 Gen 2",
    "Airdopes 141 ANC", "Rockerz 480", "Rockerz 421", "Airdopes Ace Gen 2",
    "Airdopes Alpha Gen 2", "Airdopes 213", "Airdopes 212", "Airdopes 219", "Rockerz 425",
    "Rockerz 460", "Rockerz 413", "Rockerz 412", "Rockerz 411", "Rockerz 203", "Rockerz 202",
    "Rockerz 113", "Rockerz 112", "Rockerz 111", "Rockerz Plus 550", "Rockerz 550",
    "Airdopes 138 Gen 2", "Airdopes 131 Gen 2", "Aavante Bar 950", "Aavante Bar Aspire",
    "Aavante Bar 490", "Aavante Bar 485", "Aavante Bar 480", "Airdopes 313", "Stone 358",
    "Stone 352", "Stone 350", "Stone 358 Pro", "Stone 352 Pro", "Stone 350 Pro",
    "Airdopes 301", "Airdopes 181 Pro", "Airdopes 161 ANC Elite", "Rockerz Summit",
    "Airdopes 91 Prime", "Airdopes 141 Elite ANC", "Airdopes Ace", "Airdopes 71",
    "Airdopes Ultra Plus", "Airdopes 131 Elite ANC", "Airdopes 101v2", "Airdopes Joy",
    "Rockerz 235 Pro", "Rockerz 430", "Airdopes 121 Pro Plus", "Airdopes Plus 311",
    "Airdopes 311 Pro", "Airdopes 111v2", "Rockerz 450 Pro", "Rockerz 450R", "Rockerz 450",
    "Airdopes Alpha", "Airdopes 161", "Airdopes 141", "Airdopes 131",
]

products = SDK_PRODUCTS + NON_SDK_PRODUCTS

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
    "connectivity": "Connectivity",
    "detect": "Connectivity",
    "fw update": "Firmware Update",
    "firmware update": "Firmware Update",
    "battery information": "Battery Information",
    "charging information": "Battery Information",
    "charging percentage": "Battery Information",
    "battery percentage": "Battery Information",
    "battery": "Battery Information",
    "percentage": "Battery Information",
    "noise cancellation": "Noise Cancellation Issue",
    "anc": "Noise Cancellation Issue",
    "internet": "Internet Requirement Issue",
    "offline": "Internet Requirement Issue",
    "phone number": "App Privacy Issue",
    "mobile number": "App Privacy Issue",
    "otp": "App Privacy Issue",
    "personal information": "App Privacy Issue",
    "email address": "App Privacy Issue",
    "email": "App Privacy Issue",
    "login": "Login issue",
    "log in": "Login issue",
    "log into": "Login issue",
    "log": "Login issue",
    "issue": "App Issue",
    "app open": "App Issue",
    "location": "Location Permission",
    "freez": "App Crash",
    "positive": "Positive Feedback",
    "update": "App Update issue",
    "app space": "App size",
    "download": "App download Issue",
    "install": "App download Issue",
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
    "button": "Funtionality/ Validation issue",
}

EXCEPTION_PHRASES = ["even though", "even", "though", "issue"]

#***********************************************************************#

def Positive_Reviews():
    df = pd.read_excel('Android_HearablesApp_Review.xlsx')
    positive_reviews_df = df[df['Ratings'].isin([4, 5])]
    with pd.ExcelWriter('Android_HearablesApp_Review.xlsx', engine='openpyxl', mode='a') as writer:
        positive_reviews_df.to_excel(writer, sheet_name='Positive_Reviews', index=False)
    print("Positive reviews saved to 'Android_HearablesApp_Review.xlsx' under sheet name 'Positive_Reviews'.")

def normalize_unicode_text(text):
    """Convert stylized Unicode characters to plain ASCII."""
    return unicodedata.normalize('NFKD', text).encode('ascii', 'ignore').decode('ascii')

def extract_device_names(text, threshold=70):
    if not text or not isinstance(text, str):
        return "Unknown Device"
    if len(text.split()) <= 2:
        return "Unknown Device"

    txt = text.lower()
    version_strings = re.findall(r'\b\d+\.\d+\b', txt)
    words = re.findall(r"[a-zA-Z0-9]+", txt)
    windows = words \
        + [" ".join(words[i:i+2]) for i in range(len(words)-1)] \
        + [" ".join(words[i:i+3]) for i in range(len(words)-2)]
    best_prefix = None
    best_score = 0
    for w in windows:
        result = process.extractOne(
            w, [p.lower() for p in PREFIXES], scorer=fuzz.ratio
        )
        if result and result[1] >= threshold and (
            result[1] > best_score or
            (result[1] == best_score and best_prefix and len(result[0]) > len(best_prefix))
        ):
            best_prefix = result[0]
            best_score = result[1]

    canonical_prefix = DEVICE_TYPE_ALIAS.get(best_prefix, best_prefix.title() if best_prefix else None)
    txt_no_version = re.sub(r'\b\d+\.\d+\b', '', txt)
    nums = re.findall(r"\d+", txt_no_version)

    SUFFIX_MAP = {
        "anc elite": "ANC Elite", "elite anc": "Elite ANC", "anc pro": "ANC Pro",
        "ancpro": "ANC Pro", "anc": "ANC", "pro+": "Pro+", "pro plus": "Pro Plus",
        "pro": "Pro", "plus": "Plus", "prime": "Prime", "gen 2": "Gen 2",
        "gen2": "Gen 2", "elite": "Elite", "ultra pro": "Ultra Pro", "ultra": "Ultra",
        "max": "Max", "v1": "v1", "v2": "v2", "v2 pro": "V2 Pro", "v2pro": "V2 Pro",
        "r": "R", "z plus": "Z Plus", "zplus": "Z Plus", "z": "Z", "x tws": "X TWS",
        "xtws": "X TWS", "hidef": "HiDef", "hi def": "HiDef", "loop": "loop",
        "tws": "TWS", "crown": "Crown", "ivy pro": "Ivy Pro", "ivy": "Ivy",
        "space": "Space", "lumos": "Lumos", "wave": "Wave", "neo": "Neo",
        "proclip": "ProClip", "pro clip": "ProClip", "grande": "Grande", "luxe": "Luxe",
        "iris": "Iris", "crystl": "Crystl", "lucid": "Lucid", "nebula": "Nebula",
        "eutopia": "Eutopia", "supreme": "Supreme", "aspire": "Aspire",
        "zenith pro": "Zenith Pro", "zenithpro": "Zenith Pro", "zenith": "Zenith",
        "summit": "Summit", "ace": "Ace", "101v2": "101v2", "101 v2": "101v2",
        "111v2": "111v2", "111 v2": "111v2", "joy": "Joy", "alpha": "Alpha",
        "katana": "katana", "kick": "kick", "progear": "Progear", "shiney": "Shiney",
    }

    model_num = None
    model_num_pos = None
    if nums and best_prefix:
        full_prefix_match = re.search(re.escape(best_prefix), txt_no_version)
        if full_prefix_match:
            prefix_end_pos = full_prefix_match.end()
        else:
            prefix_first_word = best_prefix.split()[0]
            prefix_match = re.search(re.escape(prefix_first_word), txt_no_version)
            prefix_end_pos = prefix_match.end() if prefix_match else 0

        suffix_start_pos = len(txt_no_version)
        for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
            m = re.search(r"\b" + re.escape(key) + r"\b", txt_no_version)
            if m and m.start() > prefix_end_pos:
                if m.start() < suffix_start_pos:
                    suffix_start_pos = m.start()

        num_positions = [(m.group(), m.start()) for m in re.finditer(r"\d+", txt_no_version)]
        between = [(n, pos) for n, pos in num_positions
                   if pos >= prefix_end_pos and pos < suffix_start_pos]
        if between:
            model_num, model_num_pos = between[0]
        else:
            if suffix_start_pos == len(txt_no_version):
                after_prefix = [(n, pos) for n, pos in num_positions if pos >= prefix_end_pos]
                if after_prefix:
                    model_num, model_num_pos = after_prefix[0]
    elif nums:
        model_num = sorted(nums, key=lambda x: (len(x), int(x)), reverse=True)[0]

    if best_prefix and "avante" in best_prefix.lower() and version_strings and model_num:
        version_str = version_strings[0]
        detected_suffix = ""
        for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
            if re.search(r"\b" + re.escape(key) + r"\b", txt):
                detected_suffix = SUFFIX_MAP[key]
                break
        device = f"{canonical_prefix} {version_str} {model_num}"
        if detected_suffix:
            device += f" {detected_suffix}"
        return device.strip()

    if best_prefix and not model_num:
        prefix_tokens = best_prefix.split()
        prefix_end = -1
        for i in range(len(words)):
            window = " ".join(words[i:i + len(prefix_tokens)])
            if fuzz.ratio(window, best_prefix) >= 95:
                prefix_end = i + len(prefix_tokens)
                break

        if prefix_end != -1:
            remaining_words = words[prefix_end:]
            window_text = " ".join(remaining_words[:2])
            detected_suffix = ""
            for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
                if re.search(r"\b" + re.escape(key) + r"\b", window_text):
                    detected_suffix = SUFFIX_MAP[key]
                    break
            parts = [canonical_prefix]
            if detected_suffix:
                parts.append(detected_suffix)
            return " ".join(parts).strip()

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
        if detected_suffix:
            return f"{model_num} {detected_suffix}".strip()
        fallback = process.extractOne(
            txt, [p.lower() for p in products], scorer=fuzz.partial_ratio
        )
        if fallback and fallback[1] >= THRESHOLD:
            matched_idx = [p.lower() for p in products].index(fallback[0])
            return products[matched_idx]
        return "Unknown Device"

    if best_prefix and model_num_pos is not None:
        tail_text = txt_no_version[model_num_pos + len(model_num):]
    else:
        prefix_end_pos = 0
        if best_prefix:
            m = re.search(re.escape(best_prefix), txt_no_version)
            if m:
                prefix_end_pos = m.end()
        tail_text = txt_no_version[prefix_end_pos:]
    tail_words = re.findall(r"[a-zA-Z0-9+]+", tail_text)
    search_region = " ".join(tail_words[:2])

    detected_suffix = ""
    for key in sorted(SUFFIX_MAP.keys(), key=lambda x: -len(x)):
        if re.search(r"\b" + re.escape(key) + r"\b", search_region):
            detected_suffix = SUFFIX_MAP[key]
            break

    device = f"{canonical_prefix} {model_num}" if canonical_prefix else "Unknown Device"
    if detected_suffix:
        device += f" {detected_suffix}"

    if device.strip() == "Unknown Device":
        result = process.extractOne(
            txt, [p.lower() for p in products], scorer=fuzz.partial_ratio
        )
        if result and result[1] >= THRESHOLD:
            matched_idx = [p.lower() for p in products].index(result[0])
            device = products[matched_idx]

    return device.strip()


def remove_emojis(text):
    emoji_pattern = re.compile(
        "["
        "\U0001F600-\U0001F64F"
        "\U0001F300-\U0001F5FF"
        "\U0001F680-\U0001F6FF"
        "\U0001F1E0-\U0001F1FF"
        "\U00002700-\U000027BF"
        "\U0001F900-\U0001F9FF"
        "\U00002600-\U000026FF"
        "]+",
        flags=re.UNICODE,
    )
    return emoji_pattern.sub("", text)


def contains_exception_phrase(text):
    text_lower = text.lower()
    for phrase in EXCEPTION_PHRASES:
        words = text_lower.split()
        for i in range(len(words)):
            window = " ".join(words[i:i+len(phrase.split())+1])
            score = fuzz.partial_ratio(phrase, window)
            if score >= THRESHOLD:
                return True
    return False


def is_valid_device(d, products_set):
    if d.isdigit():
        for product in products_set:
            nums = re.findall(r'\d+', product)
            if d in nums:
                return True
        return False
    else:
        return d in products_set


def get_sku_category(device_name):
    if not device_name or device_name == "Unknown Device":
        return "Unknown"
    d = device_name.strip()
    if d in SDK_PRODUCTS:
        return "SDK"
    if d in NON_SDK_PRODUCTS:
        return "Non-SDK"
    if d.isdigit():
        for product in SDK_PRODUCTS:
            if d in re.findall(r'\d+', product):
                return "SDK"
        for product in NON_SDK_PRODUCTS:
            if d in re.findall(r'\d+', product):
                return "Non-SDK"
    return "Unknown"


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
    for pkw in PRIORITY_KEYWORDS:
        pkw_lower = pkw.lower()
        if len(pkw_lower) <= 4:
            if re.search(r'\b' + re.escape(pkw_lower) + r'\b', review_lower):
                return pkw_lower
            continue
        if pkw_lower in review_lower:
            return pkw_lower
        if len(pkw_lower.split()) > 1:
            score = fuzz.partial_ratio(pkw_lower, review_lower)
            if score >= threshold:
                return pkw_lower

    exact_kw  = None
    exact_len = 0
    fuzzy_kw    = "Other"
    fuzzy_score = 0
    fuzzy_len   = 0

    for kw in keywords:
        kw_lower = kw.lower()
        if len(kw_lower) <= 4:
            if re.search(r'\b' + re.escape(kw_lower) + r'\b', review_lower):
                if len(kw_lower) > exact_len:
                    exact_kw  = kw_lower
                    exact_len = len(kw_lower)
            continue
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
    df = pd.read_excel('Android_HearablesApp_Review.xlsx', sheet_name='Happ_Review_Data')
    df['Reviews'] = df['Reviews'].apply(
        lambda x: normalize_unicode_text(str(x)).strip() if pd.notna(x) else x
    )
    negative_reviews_df = df[df['Ratings'].isin([1,2,3])].copy()
    sentiment_analyzer = pipeline("sentiment-analysis")

    reasons, DUT = [], []
    for review in negative_reviews_df['Reviews']:
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

    consolidated_reasons = []
    products_set = set(products)

    SDK_ONLY_FEATURE_REASONS = {"anc", "eq", "equalizer", "custom equalizer", "noise cancellation"}
    BARE_PRODUCT_SERIES = {
        "Nirvana", "Airdopes", "Stone", "Rockerz", "Aavante", "Aavante Bar", "Immortal",
    }

    for review, reason, dut in zip(negative_reviews_df['Reviews'].tolist(), reasons, DUT):
        r = reason.lower()
        d = dut.strip()
        review_lower = str(review).lower() if pd.notna(review) else ""
        sku_category = get_sku_category(d)

        if d != "Unknown Device" and d not in BARE_PRODUCT_SERIES and not is_valid_device(d, products_set):
            consolidated = "SKU not supported"
        elif sku_category == "Non-SDK" and r in SDK_ONLY_FEATURE_REASONS:
            consolidated = "Non-SDK less feature"
        elif r == "update":
            if any(sig in review_lower for sig in FORCE_UPDATE_SIGNALS):
                consolidated = "Force App Update Not Required"
            else:
                consolidated = "App Update issue"
        else:
            base_consolidated = CONSOLIDATION_MAP.get(r, "Generic Comments")
            if base_consolidated in ["Equalizer issue", "Noise Cancellation Issue"]:
                consolidated = "Equalizer and Anc removed for Airhoa SDK to support 16kb page"
            else:
                consolidated = base_consolidated

        consolidated_reasons.append(consolidated)

    negative_reviews_df['Consolidated Reason'] = consolidated_reasons

    with pd.ExcelWriter('Android_HearablesApp_Review.xlsx', engine='openpyxl', mode='a') as writer:
        negative_reviews_df.to_excel(writer, sheet_name='Negative_Reviews', index=False)

    print("Negative reviews and reasons saved to 'Negative_Reviews' sheet.")
    print("✅ Equalizer and Anc removed for Airhoa SDK to support 16kb page")


def SheetNameChange():
    directory_path = os.path.dirname(os.path.realpath(__file__))
    file_path = 'Android_HearablesApp_Review.xlsx'
    current_sheet_name = 'Sheet1'
    new_sheet_name = 'Happ_Review_Data'
    try:
        wb = openpyxl.load_workbook(file_path)
        if current_sheet_name in wb.sheetnames:
            sheet = wb[current_sheet_name]
            sheet.title = new_sheet_name
            wb.save(file_path)
            print(f"The sheet name '{current_sheet_name}' has been changed to '{new_sheet_name}'.")
        else:
            print(f"The sheet '{current_sheet_name}' does not exist in the Excel file.")
        wb.close()
    except FileNotFoundError:
        print(f"The file '{file_path}' does not exist.")
    except Exception as e:
        print(f"An error occurred: {str(e)}")


def Sorting_Xlsx(sheet_name):
    workbook = openpyxl.load_workbook('Android_HearablesApp_Review.xlsx')
    worksheet = workbook[sheet_name]
    column_widths = {'A': 15, 'B': 15, 'C': 8, 'D': 42, 'E': 12.5, 'F': 11, 'G': 11, 'H': 10, 'I': 10, 'J': 48, 'K': 15, 'L': 21}
    for column, width in column_widths.items():
        worksheet.column_dimensions[column].width = width
    for row in worksheet.iter_rows():
        for cell in row:
            cell.alignment = openpyxl.styles.Alignment(wrap_text=True, vertical="center")
    workbook.save('Android_HearablesApp_Review.xlsx')
    workbook.close()
    print(f"Xlsx modification done for sheet '{sheet_name}'.")


def Review_Date_Calculator():
    global lowest_date
    global latest_date
    df = pd.read_excel('Android_HearablesApp_Review.xlsx')
    df['Date Time'] = pd.to_datetime(df['Date Time'])
    df['DateOnly'] = df['Date Time'].dt.date
    lowest_date = df['DateOnly'].min()
    latest_date = df['DateOnly'].max()
    print("Lowest Date in the DateTimeColumn:", lowest_date)
    print("Latest Date in the DateTimeColumn:", latest_date)


def Negative_Review_Analysis():
    df = pd.read_excel('Android_HearablesApp_Review.xlsx', sheet_name='Negative_Reviews')
    reason_counts = df['Consolidated Reason'].value_counts()
    print("Negative Review Analysis:")
    for reason, count in reason_counts.items():
        print(f"{reason}: {count}")
    if 'DUT' in df.columns:
        sku_counts = df['DUT'].apply(get_sku_category).value_counts()
        print("\nNegative Reviews by SKU Category:")
        for category, count in sku_counts.items():
            print(f"{category}: {count}")
    return reason_counts


def Summary():
    global App_Rating
    global App_Version
    global lowest_date
    global latest_date

    print("Generating styled Summary sheet...")

    df = pd.read_excel('Android_HearablesApp_Review.xlsx')
    pos_count = len(df[df['Ratings'].isin([4, 5])])
    neg_count = len(df[df['Ratings'].isin([1, 2, 3])])

    app_name = "Android Hearables App " + str(App_Version)
    app_rating = App_Rating + '*'
    review_data = str(lowest_date) + ' to ' + str(latest_date)

    app_details_data = {
        'App Name': [app_name],
        'Latest App Version': [App_Version],
        'Review Date': [review_data]
    }
    app_df = pd.DataFrame(app_details_data)
    counts_df = pd.DataFrame({
        'Category': ['Positive Reviews (4 & 5)', 'Negative Reviews (<=3)'],
        'Count': [pos_count, neg_count]
    })

    neg_df = pd.read_excel('Android_HearablesApp_Review.xlsx', sheet_name='Negative_Reviews')
    reason_counts = neg_df['Consolidated Reason'].value_counts()
    summary_df = pd.DataFrame({
        'Summary (-ve reviews)': reason_counts.index,
        'Count': reason_counts.values
    })

    def make_multiline(df_filter):
        df_filter = df_filter[df_filter['DUT'] != 'Unknown Device']
        if df_filter.empty:
            return "No valid devices found"
        counts = df_filter['DUT'].value_counts()
        return "\n".join([f"{dev} - {cnt}" for dev, cnt in counts.items()])

    summary_df['Remarks'] = [
        make_multiline(neg_df[neg_df['Consolidated Reason'] == reason])
        for reason in summary_df['Summary (-ve reviews)']
    ]

    if 'Positive Feedback' in summary_df['Summary (-ve reviews)'].values:
        summary_df.loc[
            summary_df['Summary (-ve reviews)'] == 'Positive Feedback', 'Remarks'
        ] = "Reviews are good but ratings are <= 3"

    if 'DUT' in neg_df.columns:
        sku_cat_counts = neg_df['DUT'].apply(get_sku_category).value_counts()
        sku_df = pd.DataFrame({
            'SKU Category (-ve reviews)': sku_cat_counts.index,
            'Count': sku_cat_counts.values
        })
    else:
        sku_df = pd.DataFrame()

    wb = load_workbook('Android_HearablesApp_Review.xlsx')
    ws = wb['Summary'] if 'Summary' in wb.sheetnames else wb.create_sheet('Summary')
    ws.delete_rows(1, ws.max_row)

    ws.append(["📊 Review Summary Report"])
    ws.merge_cells('A1:C1')
    title = ws['A1']
    title.font = Font(size=14, bold=True, color="FFFFFF")
    title.alignment = Alignment(horizontal='center', vertical='center')
    title.fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ws.append([])

    for df_part in [app_df, pd.DataFrame(), counts_df, pd.DataFrame(), sku_df, pd.DataFrame(), summary_df]:
        for r in dataframe_to_rows(df_part, index=False):
            ws.append(r)

    thin = Side(border_style="thin", color="000000")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)
    wrap = Alignment(wrap_text=True, vertical='top')
    blue_header = PatternFill(start_color="A7C7E7", end_color="A7C7E7", fill_type="solid")
    green_section = PatternFill(start_color="C8E6C9", end_color="C8E6C9", fill_type="solid")

    col_widths = {'A': 25, 'B': 15, 'C': 35, 'D': 23}
    for col, width in col_widths.items():
        ws.column_dimensions[col].width = width

    for row in ws.iter_rows():
        for cell in row:
            cell.border = border
            cell.alignment = wrap

    for row in ws.iter_rows(min_row=1, max_col=3):
        values = [str(c.value).strip().lower() if c.value else "" for c in row]
        if "category" in values or "summary (-ve reviews)" in values or "sku category (-ve reviews)" in values:
            for cell in row:
                cell.fill = blue_header
                cell.font = Font(bold=True)
        elif "app name" in values:
            for cell in row:
                cell.fill = green_section
                cell.font = Font(bold=True)

    for row in ws.iter_rows():
        for cell in row:
            if hasattr(cell, "column_letter") and cell.column_letter == 'C':
                cell.alignment = Alignment(wrap_text=True, horizontal='center', vertical='center')

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
    try:
        workbook = openpyxl.load_workbook(file_path)
        workbook.close()
        print("Closed the Excel file.")
    except Exception as e:
        print(f"Error closing the Excel file: {e}")
