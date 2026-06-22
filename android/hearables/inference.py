"""
inference.py — Downloads latest BERT model from Google Drive using gdown,
classifies Negative_Reviews and writes AI_Confidence + AI_Consolidated_Reason
back to the Negative_Reviews sheet.
"""

import os
import re
import shutil
import unicodedata
import warnings
import pandas as pd
import torch
import gdown
from transformers import BertTokenizerFast, BertForSequenceClassification
from rapidfuzz import fuzz, process
from openpyxl import load_workbook

warnings.filterwarnings("ignore")

# ── CONFIG ─────────────────────────────────────────────────────────────
EXCEL_PATH   = "Android_HearablesApp_Review.xlsx"
DRIVE_FOLDER = os.environ.get('DRIVE_FOLDER')
if not DRIVE_FOLDER:
    raise SystemExit("❌ Missing required secret: DRIVE_FOLDER")
MODEL_CACHE  = "model_cache"
OUTPUT_SHEET = "Negative_Reviews"
BATCH_SIZE   = 16
MAX_LEN      = 128

# ── PRODUCTS ────────────────────────────────────────────────────────────
DEVICE_TYPES = [
    "Stone", "Rockerz", "Airdopes Atom", "Airdopes Prime", "Airdopes",
    "Nirvana", "Party Pal", "immortal", "avante bar", "aavante", "party pal", "unity",
]
DEVICE_TYPE_ALIAS = {
    "aavante": "Aavante Bar", "avante bar": "Aavante Bar",
    "unity": "Boat Unity", "airpods": "Airdopes",
}
THRESHOLD = 80
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
    "Airdopes 101 v2","Airdopes Joy","Rockerz 235 Pro","Rockerz 430",
    "Airdopes 121 Pro Plus","Airdopes Plus 311","Airdopes 311 Pro","Airdopes 111v2",
    "Rockerz Trinity Gen 2","Airdopes Prime 701 ANC","Rockerz 450 Pro","Rockerz 450R",
    "Rockerz 450","Nirvana Ivy Pro","Nirvana Zenith Pro","Nirvana Iris","PartyPal 600",
    "Rockerz 650 Pro","Rockerz 551 ANC Pro","Nirvana X TWS","Airdopes 800 HiDef",
    "Nirvana Crystl","Airdopes Ultra Pro","Nirvana Lucid","Airdopes Alpha",
    "Airdopes 161","Airdopes 141","Airdopes 131","Nirvana Ion ANC Pro",
    "Rockerz 210 ANC","Nirvana Ivy","Rockerz 255 Z Plus","Stone Lumos","Nirvana Space",
    "NIRVANA ZENITH","NIRVANA NEBULA","Nirvana Eutopia","Airdopes Supreme",
    "Rockerz 255 ANC","Nirvana Ion ANC","Nirvana Ion","Airdopes Flex 454 ANC",
    "Airdopes 800","Rockerz 255 Max","Airdopes 300","NIRVANA 525ANC","Airdopes 341ANC",
    "Airdopes 393ANC","Airdopes 172","Rockerz 255 Pro+","Rockerz 333 Pro","Rockerz 333",
    "Rockerz 330 Pro","Nirvana Crown","Rockerz 202","Airdopes Prime 700 ANC","Party Pal 600",
]

FORCE_UPDATE_SIGNALS = [
    "without updating","without update","let me use","skip update","force update",
    "forced to update","dont want to update","don't want to update",
    "not updating","without upgrading",
]

# ── HELPER FUNCTIONS ────────────────────────────────────────────────────
def normalize_unicode_text(text):
    return unicodedata.normalize("NFKD", text).encode("ascii", "ignore").decode("ascii")

def is_valid_device(d, products_set):
    if d.isdigit():
        return any(d in re.findall(r"\d+", p) for p in products_set)
    return d in products_set

def extract_device_names(text, threshold=70):
    if not text or not isinstance(text, str): return "Unknown Device"
    if len(text.split()) <= 2: return "Unknown Device"
    txt = text.lower()
    words = re.findall(r"[a-zA-Z0-9]+", txt)
    windows = words + [" ".join(words[i:i+2]) for i in range(len(words)-1)]
    best_prefix, best_score = None, 0
    for w in windows:
        result = process.extractOne(w, [p.lower() for p in DEVICE_TYPES], scorer=fuzz.ratio)
        if result and result[1] >= threshold and result[1] > best_score:
            best_prefix, best_score = result[0], result[1]
    if not best_prefix:
        fallback = process.extractOne(txt, [p.lower() for p in products], scorer=fuzz.partial_ratio)
        if fallback and fallback[1] >= THRESHOLD:
            return products[[p.lower() for p in products].index(fallback[0])]
        return "Unknown Device"
    canonical = DEVICE_TYPE_ALIAS.get(best_prefix, best_prefix.title())
    txt_nv = re.sub(r"\b\d+\.\d+\b", "", txt)
    nums = re.findall(r"\d+", txt_nv)
    model_num = nums[0] if nums else None
    SUFFIX_MAP = {
        "anc elite":"ANC Elite","anc pro":"ANC Pro","anc":"ANC","pro+":"Pro+",
        "pro plus":"Pro+","pro":"Pro","plus":"Plus","prime":"Prime",
        "gen 2":"Gen 2","gen2":"Gen 2","elite":"Elite","ultra pro":"Ultra Pro",
        "ultra":"Ultra","max":"Max","v2":"v2","r":"R","z plus":"Z Plus","z":"Z",
        "tws":"TWS","hidef":"HiDef","ion anc pro":"Ion ANC Pro","ion anc":"Ion ANC",
        "ion":"Ion","ivy pro":"Ivy Pro","ivy":"Ivy","crown":"Crown",
        "elite anc":"Elite ANC","eliteanc":"Elite ANC"
    }
    detected_suffix = ""
    for key in sorted(SUFFIX_MAP, key=lambda x: -len(x)):
        if re.search(r"\b" + re.escape(key) + r"\b", txt):
            detected_suffix = SUFFIX_MAP[key]; break
    device = f"{canonical} {model_num}" if model_num else canonical
    if detected_suffix: device += f" {detected_suffix}"
    return device.strip()

# ── STEP 1: GET MODEL FROM GOOGLE DRIVE ─────────────────────────────────
def download_latest_model():
    print("Checking Google Drive for latest model...")

    os.makedirs(MODEL_CACHE, exist_ok=True)

    # Check if model already cached locally
    cached_folders = sorted(
        [i for i in os.listdir(MODEL_CACHE)
         if i.startswith("RA_model") and os.path.isdir(os.path.join(MODEL_CACHE, i))],
        reverse=True
    )

    if cached_folders:
        cached_model = os.path.join(MODEL_CACHE, cached_folders[0])
        if os.path.exists(os.path.join(cached_model, "model.safetensors")):
            print(f"Using cached model: {cached_folders[0]}")
            return cached_model

    # Download from Google Drive if not cached
    print("Downloading model from Google Drive...")
    url = f"https://drive.google.com/drive/folders/{DRIVE_FOLDER}"

    for attempt in range(3):
        try:
            print(f"Attempt {attempt + 1}/3...")
            gdown.download_folder(
                url, output=MODEL_CACHE, quiet=False, use_cookies=False
            )
            break
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt == 2:
                raise Exception("Failed to download model after 3 attempts!")

    folders = sorted(
        [i for i in os.listdir(MODEL_CACHE)
         if i.startswith("RA_model") and os.path.isdir(os.path.join(MODEL_CACHE, i))],
        reverse=True
    )

    if not folders:
        raise Exception("No model folder found after download!")

    latest = os.path.join(MODEL_CACHE, folders[0])
    print(f"Using model: {folders[0]}")
    return latest

# ── STEP 2: RUN INFERENCE ───────────────────────────────────────────────
def run_inference(model_path):
    print("=" * 60)
    print(f"Loading model from: {model_path}")
    print("=" * 60)

    tokenizer = BertTokenizerFast.from_pretrained(model_path)
    model     = BertForSequenceClassification.from_pretrained(model_path)
    model.eval()

    device   = torch.device("cuda" if torch.cuda.is_available() else "cpu")
    model.to(device)
    id2label = model.config.id2label

    print(f"  Device  : {device}")
    print(f"  Classes : {len(id2label)}")

    # Load Negative_Reviews sheet
    df = pd.read_excel(EXCEL_PATH, sheet_name="Negative_Reviews")
    df["Reviews"] = df["Reviews"].apply(
        lambda x: normalize_unicode_text(str(x)).strip() if pd.notna(x) else ""
    )
    texts = df["Reviews"].tolist()
    print(f"  Reviews to classify: {len(texts)}")

    # Batch inference
    print("  Running model inference...")
    all_labels, all_conf = [], []

    for i in range(0, len(texts), BATCH_SIZE):
        batch_texts = texts[i:i + BATCH_SIZE]
        enc = tokenizer(
            batch_texts, max_length=MAX_LEN,
            padding=True, truncation=True, return_tensors="pt",
        )
        with torch.no_grad():
            logits = model(
                input_ids=enc["input_ids"].to(device),
                attention_mask=enc["attention_mask"].to(device),
            ).logits
        probs = torch.softmax(logits, dim=-1).cpu().numpy()
        preds = probs.argmax(axis=-1)
        for prob_row, pred_idx in zip(probs, preds):
            all_labels.append(id2label[int(pred_idx)])
            all_conf.append(float(prob_row[pred_idx]))

    print("  Running DUT extraction...")
    duts = [extract_device_names(t) for t in texts]

    print("  Building consolidated reasons...")
    products_set      = set(products)
    consolidated_list = []

    for review, label, dut in zip(texts, all_labels, duts):
        review_lower = str(review).lower()
        if dut != "Unknown Device" and not is_valid_device(dut, products_set):
            consolidated = "SKU not supported"
        elif label.lower() in ("app update issue", "app update"):
            if any(sig in review_lower for sig in FORCE_UPDATE_SIGNALS):
                consolidated = "Force App Update Not Required"
            else:
                consolidated = label
        else:
            consolidated = label
        consolidated_list.append(consolidated)

    print("  Adding QA Comments column (empty, for manual QA input)...")

    # Add columns to dataframe
    df["AI_Confidence\n[Testing Purpose]"]          = [f"{c:.1%}" for c in all_conf]
    df["AI_Consolidated_Reason\n[Testing Purpose]"] = consolidated_list
    df["QA Comments\n[give the comments based\non Consolidated reason]"] = ""

    # Write back to Negative_Reviews sheet
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        df.to_excel(writer, sheet_name=OUTPUT_SHEET, index=False)

    # Enable wrap text on header row so the multi-line labels display correctly
    from openpyxl.styles import Alignment
    wb_fmt = load_workbook(EXCEL_PATH)
    ws_fmt = wb_fmt[OUTPUT_SHEET]
    for cell in ws_fmt[1]:
        cell.alignment = Alignment(wrap_text=True, vertical="center", horizontal="center")
    wb_fmt.save(EXCEL_PATH)

    print(f"\nResults written to sheet '{OUTPUT_SHEET}' in '{EXCEL_PATH}'")
    print("Columns added: AI_Confidence [Testing Purpose] | AI_Consolidated_Reason [Testing Purpose] | QA Comments (empty)")

# ── MAIN ────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    model_path = download_latest_model()
    run_inference(model_path)
