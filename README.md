# boAt Review Analyzer — All Platforms (Consolidated)

Runs all 6 weekly review pipelines — Android Crest, Hearables, Wearables and
iOS Crest, Hearables, Wearables — one after another in a single GitHub
Actions run, then sends **one** consolidated email with all generated Excel
files attached. Apps with no reviews that week are listed as a note in the
email instead of an attachment.

## Why this structure

Each pipeline's files are kept exactly as they are in their own individual
repo — nothing has been merged or rewritten. If you ever need to fix or
adjust one specific app's logic, go directly into its folder and edit only
that file. The orchestrator (`run_all.py`) just calls each one's `main.py`
(and `inference.py` where applicable) as a separate subprocess, in its own
working directory, then collects whatever Excel files come out.

## Folder structure

```
boat-review-all/
├── android/
│   ├── crest/
│   │   ├── main.py
│   │   └── BFAT_CrestApp_PlayStore_Code.py
│   ├── hearables/
│   │   ├── main.py
│   │   ├── inference.py          ← BERT classification step, runs after main.py
│   │   ├── BFAT_HApp_PlayStore_Code.py
│   │   └── uploadtodrive.py       (unused leftover, kept as-is)
│   └── wearables/
│       ├── main.py
│       └── BFAT_WearablesApp_PlayStore_Code.py
├── ios/
│   ├── crest/main.py
│   ├── hearables/main.py
│   └── wearables/main.py
├── run_all.py                     ← orchestrator, entry point
├── requirements.txt
├── trigger_all_reviews.gs         ← paste into Apps Script
└── .github/workflows/analyze.yml
```

## GitHub Secrets required

| Secret | Used by | Value |
|---|---|---|
| `GOOGLE_PLAY_CREDENTIALS_JSON_CREST` | Android Crest | Contents of `cove-boat-*.json` service account key |
| `GOOGLE_PLAY_CREDENTIALS_JSON_HEARABLE_WEARABLE` | Android Hearables + Wearables | Contents of `playstoreautomation-*.json` service account key (shared) |
| `ANDROID_APP_ID_CREST` | Android Crest | `com.coveiot.android.boat` |
| `ANDROID_APP_ID_HEARABLES` | Android Hearables | `com.boAt.hearables` |
| `ANDROID_APP_ID_WEARABLES` | Android Wearables | `com.boAt.wristgear` |
| `APPLE_PRIVATE_KEY_CREST` | iOS Crest | Contents of CrestApp `.p8` key |
| `APPLE_PRIVATE_KEY_HEARABLE_WEARABLE` | iOS Hearables + Wearables | Contents of shared `.p8` key |
| `APPLE_KEY_ID_CREST` | iOS Crest | `8BQP29JYK7` |
| `APPLE_ISSUER_ID_CREST` | iOS Crest | `69a6de8f-b92a-47e3-e053-5b8c7c11a4d1` |
| `APPLE_APP_ID_CREST` | iOS Crest | `1573391983` |
| `APP_DISPLAY_NAME_IOS_CREST` | iOS Crest | `iOS Crest App` |
| `APPLE_KEY_ID_HEARABLE_WEARABLE` | iOS Hearables + Wearables | `JU92BX437A` |
| `APPLE_ISSUER_ID_HEARABLE_WEARABLE` | iOS Hearables + Wearables | `fdafa4b5-2cef-4f84-b7b7-c72e03f67a8e` |
| `APPLE_APP_ID_HEARABLES` | iOS Hearables | `1592550875` |
| `APP_DISPLAY_NAME_IOS_HEARABLES` | iOS Hearables | `iOS Hearables App` |
| `APPLE_APP_ID_WEARABLES` | iOS Wearables | `1542443145` |
| `APP_DISPLAY_NAME_IOS_WEARABLES` | iOS Wearables | `iOS Wearables App` |
| `HF_TOKEN` | Android Hearables BERT step | Not currently used by any code — safe to omit |
| `EMAIL_USERNAME` | All | Sender Gmail address |
| `EMAIL_PASSWORD` | All | Gmail **App Password** |
| `NOTIFY_EMAIL` | Optional | Recipient — defaults to `boatqatest@gmail.com` if not set |

All scripts now require these secrets to be set — there is no hardcoded
fallback value in the code anymore. If a required secret is missing, the
script will exit immediately with a clear error message naming which
secret is missing, rather than silently using a stale or incorrect default.

## How run_all.py works

1. Loops through all 6 pipelines in order: Android Crest → Android Hearables
   → Android Wearables → iOS Crest → iOS Hearables → iOS Wearables.
2. For each one, runs its step(s) (`main.py`, then `inference.py` for
   Hearables Android only) as a subprocess inside that app's own folder —
   so relative file paths and imports inside each script work exactly as
   they did in their original standalone repo.
3. Injects the correct credentials into the exact environment variable name
   each script expects (e.g. both Android scripts read
   `GOOGLE_PLAY_CREDENTIALS_JSON`, so the orchestrator sets that variable to
   the Crest-specific or shared value depending on which pipeline is
   currently running).
4. After each subprocess finishes, checks whether the expected `.xlsx` was
   created. If not, that app is treated as "no reviews this week" rather
   than a hard failure.
5. Once all 6 have run, sends one email: successful apps get their Excel
   attached; apps with no data get a line like
   `NOTE: No Wearables iOS review for this week`.
6. All generated Excel files are also uploaded as a GitHub Artifact
   (90-day retention) regardless of email outcome.

## Apps Script setup

1. Paste `trigger_all_reviews.gs` into your Apps Script project.
2. Update `YOUR_NEW_GITHUB_PAT_HERE` with a token that has `repo` + `workflow` scope.
3. Update `GITHUB_REPO` if your repo name differs from `boat-review-all`.
4. Run once: `setGitHubTokenAllReviews`
5. Run once: `createMondayTriggerAllReviews` — sets the weekly Monday 11 AM trigger.

## Manual test

GitHub → Actions → "boAt Weekly Review Analyzer (All Platforms)" → Run workflow.

## Security cleanup performed

Two things were cleaned up from the original individual repos before
consolidating here:

1. The original `boat-review-crest-andriod` repo had its service account
   JSON key file committed directly in the repo root. That file was **not**
   copied into this consolidated repo — only the secret-based approach is
   used here. Consider removing that committed key file from the original
   repo's git history if it hasn't been already.

2. All three Android BFAT helper files (`BFAT_CrestApp_PlayStore_Code.py`,
   `BFAT_HApp_PlayStore_Code.py`, `BFAT_WearablesApp_PlayStore_Code.py`)
   originally contained large, unused legacy Selenium web-scraping functions
   left over from before the API migration. These functions contained
   hardcoded internal URLs with real Play Console developer IDs and app IDs,
   and in one case a Zoho/Google account login URL. Since `main.py` no
   longer calls any of these functions, they have been removed entirely —
   each BFAT file now contains only the post-processing functions actually
   used (Excel formatting, keyword classification, Summary generation).
   `selenium`, `undetected-chromedriver`, and `beautifulsoup4` were removed
   from `requirements.txt` accordingly since nothing imports them anymore.

3. Previously hardcoded `KEY_ID`, `ISSUER_ID`, `APP_ID`, and `APP_NAME`
   values in the iOS scripts, and `APP_ID` in the Android scripts, are now
   read from environment variables (injected by `run_all.py` from GitHub
   Secrets) with the original values kept only as local-fallback defaults
   for convenience when testing outside GitHub Actions.
