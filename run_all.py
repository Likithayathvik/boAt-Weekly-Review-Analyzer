"""
run_all.py — Orchestrator for all 6 boAt Review Analyzers
Runs Android (Crest, Hearables, Wearables) and iOS (Crest, Hearables, Wearables)
one by one as completely separate, untouched scripts. Each pipeline's own
main.py / inference.py / send_email.py are never modified or merged — they
just get called in sequence from their own folder.

After all 6 finish, this script collects whichever Excel files were
generated and sends ONE consolidated email with all of them attached.
If a particular app had no reviews this week, that app is listed in the
email body as a note instead of being attached.
"""

import os
import sys
import subprocess
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.application import MIMEApplication
from datetime import datetime

BASE_DIR = os.path.dirname(os.path.abspath(__file__))

# ── PIPELINE DEFINITIONS ──────────────────────────────────────────────────────
# Each entry:
#   label          — friendly name shown in email/logs
#   folder         — relative folder containing this pipeline's files
#   steps          — list of script filenames to run IN ORDER inside that folder
#                    (most apps have just ["main.py"]; Hearables Android also
#                     runs inference.py afterward for BERT classification)
#   xlsx_name      — expected output Excel filename
#   apple_key_env  — env var name holding the Apple private key for this app
#                    (None for Android pipelines)
PIPELINES = [
    {
        "label": "Android - CrestApp",
        "folder": "android/crest",
        "steps": ["main.py"],
        "xlsx_name": "Android_CrestApp_Review.xlsx",
        "apple_key_env": None,
        "extra_env": {
            "GOOGLE_PLAY_CREDENTIALS_JSON": "GOOGLE_PLAY_CREDENTIALS_JSON_CREST",
            "ANDROID_APP_ID": "ANDROID_APP_ID_CREST",
        },
    },
    {
        "label": "Android - HearablesApp",
        "folder": "android/hearables",
        "steps": ["main.py", "inference.py"],
        "xlsx_name": "Android_HearablesApp_Review.xlsx",
        "apple_key_env": None,
        "extra_env": {
            "GOOGLE_PLAY_CREDENTIALS_JSON": "GOOGLE_PLAY_CREDENTIALS_JSON_HEARABLE_WEARABLE",
            "ANDROID_APP_ID": "ANDROID_APP_ID_HEARABLES",
        },
    },
    {
        "label": "Android - WearablesApp",
        "folder": "android/wearables",
        "steps": ["main.py"],
        "xlsx_name": "Android_WearablesApp_Review.xlsx",
        "apple_key_env": None,
        "extra_env": {
            "GOOGLE_PLAY_CREDENTIALS_JSON": "GOOGLE_PLAY_CREDENTIALS_JSON_HEARABLE_WEARABLE",
            "ANDROID_APP_ID": "ANDROID_APP_ID_WEARABLES",
        },
    },
    {
        "label": "iOS - CrestApp",
        "folder": "ios/crest",
        "steps": ["main.py"],
        "xlsx_name": "iOS_CrestApp_Review.xlsx",
        "apple_key_env": "APPLE_PRIVATE_KEY_CREST",
        "extra_env": {
            "APPLE_KEY_ID":     "APPLE_KEY_ID_CREST",
            "APPLE_ISSUER_ID":  "APPLE_ISSUER_ID_CREST",
            "APPLE_APP_ID":     "APPLE_APP_ID_CREST",
            "APP_DISPLAY_NAME": "APP_DISPLAY_NAME_IOS_CREST",
        },
    },
    {
        "label": "iOS - HearablesApp",
        "folder": "ios/hearables",
        "steps": ["main.py"],
        "xlsx_name": "iOS_HearablesApp_Review.xlsx",
        "apple_key_env": "APPLE_PRIVATE_KEY_HEARABLE_WEARABLE",
        "extra_env": {
            "APPLE_KEY_ID":     "APPLE_KEY_ID_HEARABLE_WEARABLE",
            "APPLE_ISSUER_ID":  "APPLE_ISSUER_ID_HEARABLE_WEARABLE",
            "APPLE_APP_ID":     "APPLE_APP_ID_HEARABLES",
            "APP_DISPLAY_NAME": "APP_DISPLAY_NAME_IOS_HEARABLES",
        },
    },
    {
        "label": "iOS - WearablesApp",
        "folder": "ios/wearables",
        "steps": ["main.py"],
        "xlsx_name": "iOS_WearablesApp_Review.xlsx",
        "apple_key_env": "APPLE_PRIVATE_KEY_HEARABLE_WEARABLE",
        "extra_env": {
            "APPLE_KEY_ID":     "APPLE_KEY_ID_HEARABLE_WEARABLE",
            "APPLE_ISSUER_ID":  "APPLE_ISSUER_ID_HEARABLE_WEARABLE",
            "APPLE_APP_ID":     "APPLE_APP_ID_WEARABLES",
            "APP_DISPLAY_NAME": "APP_DISPLAY_NAME_IOS_WEARABLES",
        },
    },
]

# ── RUN ONE PIPELINE (possibly multiple steps) ────────────────────────────────
def run_pipeline(pipeline):
    label   = pipeline["label"]
    folder  = pipeline["folder"]
    steps   = pipeline["steps"]
    xlsx    = pipeline["xlsx_name"]

    print("\n" + "=" * 70)
    print(f"RUNNING: {label}")
    print("=" * 70)

    workdir = os.path.join(BASE_DIR, folder)

    # Build environment for this pipeline's subprocess(es)
    env = os.environ.copy()
    if pipeline.get("apple_key_env"):
        env["APPLE_PRIVATE_KEY"] = os.environ.get(pipeline["apple_key_env"], "")
    for target_var, source_var in pipeline.get("extra_env", {}).items():
        env[target_var] = os.environ.get(source_var, "")

    for step_script in steps:
        print(f"\n  → Step: {step_script}")
        try:
            result = subprocess.run(
                [sys.executable, step_script],
                cwd=workdir,
                capture_output=True,
                text=True,
                timeout=900,
                env=env,
            )
            print(result.stdout[-3000:])
            if result.returncode != 0:
                print(f"  ❌ {step_script} FAILED (exit code {result.returncode})")
                print(result.stderr[-2000:])
                return None
        except subprocess.TimeoutExpired:
            print(f"  ❌ {step_script} TIMED OUT after 15 minutes")
            return None
        except Exception as e:
            print(f"  ❌ {step_script} ERROR: {e}")
            return None

    xlsx_path = os.path.join(workdir, xlsx)
    if os.path.exists(xlsx_path):
        print(f"\n✅ {label} — Excel generated: {xlsx}")
        return xlsx_path
    else:
        print(f"\n⚠️  {label} — completed but no Excel found (likely no reviews this week)")
        return None

# ── SEND CONSOLIDATED EMAIL ───────────────────────────────────────────────────
def send_consolidated_email(results):
    EMAIL_USERNAME = os.environ['EMAIL_USERNAME']
    EMAIL_PASSWORD = os.environ['EMAIL_PASSWORD']
    TO_EMAIL       = os.environ.get('NOTIFY_EMAIL', '').strip() or 'boatqatest@gmail.com'

    # CC recipients — read from CC_EMAILS secret (comma-separated).
    # Currently commented out below; uncomment the msg['Cc'] line and the
    # recipients list line in smtplib.send_message() to activate.
    CC_EMAILS = os.environ.get('CC_EMAILS', '').strip()
    cc_list   = [e.strip() for e in CC_EMAILS.split(',') if e.strip()] if CC_EMAILS else []

    succeeded = [(label, path) for label, path in results if path]
    no_data   = [label for label, path in results if not path]

    msg            = MIMEMultipart()
    msg['From']    = EMAIL_USERNAME
    msg['To']      = TO_EMAIL
    msg['Cc']      = ", ".join(cc_list)
    msg['Subject'] = f"boAt Weekly Review Reports (All Platforms) — {datetime.now().strftime('%d %b %Y')}"

    lines = ["Hello,", "", "Please find attached this week's review reports.", ""]

    if succeeded:
        lines.append(f"Attached ({len(succeeded)}/6):")
        for label, _ in succeeded:
            lines.append(f"  - {label}")
        lines.append("")

    if no_data:
        for label in no_data:
            short_name = label.replace("Android - ", "").replace("iOS - ", "")
            platform   = "Android" if label.startswith("Android") else "iOS"
            lines.append(f"NOTE: No {short_name.replace('App','')} {platform} review for this week")
        lines.append("")

    lines.append("— boAt R&D QA Automation")

    msg.attach(MIMEText("\n".join(lines), 'plain'))

    for label, path in succeeded:
        filename = os.path.basename(path)
        with open(path, 'rb') as f:
            attachment = MIMEApplication(f.read(), _subtype='xlsx')
            attachment.add_header(
                'Content-Disposition', 'attachment',
                filename=f"{os.path.splitext(filename)[0]}_{datetime.now().strftime('%Y-%m-%d')}.xlsx"
            )
            msg.attach(attachment)

    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as server:
        server.login(EMAIL_USERNAME, EMAIL_PASSWORD)
        server.send_message(msg, to_addrs=[TO_EMAIL] + cc_list)

    print(f"\n✅ Consolidated email sent to {TO_EMAIL} with {len(succeeded)} attachment(s)")
    if cc_list:
        print(f"   CC: {', '.join(cc_list)}")

# ── MAIN ─────────────────────────────────────────────────────────────────────
if __name__ == "__main__":
    print("#" * 70)
    print("# boAt Weekly Review Analyzer — Running ALL 6 pipelines")
    print("#" * 70)

    results = []
    for pipeline in PIPELINES:
        xlsx_path = run_pipeline(pipeline)
        results.append((pipeline["label"], xlsx_path))

    print("\n" + "#" * 70)
    print("# SUMMARY")
    print("#" * 70)
    for label, path in results:
        status = "✅ SUCCESS" if path else "⚠️  NO DATA / FAILED"
        print(f"  {status:20s} — {label}")

    send_consolidated_email(results)

    print("\n" + "#" * 70)
    print("# ALL PIPELINES COMPLETE")
    print("#" * 70)
