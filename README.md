# Jev Test Runner

> **It's 2026. You don't need a test automation framework.**
> Write test cases in Excel. The AI clicks, types and verifies them in a real browser. PASS / FAIL is written back into the sheet.

Jev Test Runner takes **manual test cases written in plain English in an Excel sheet** and runs them in Chrome. You don't write Selenium scripts, locators or page objects.

- **[Playwright](https://playwright.dev/python/)** drives the browser.
- **TypeSafe Jev** reads each step, chooses the action, the on-screen element and the test data value, and then checks the Expected Result against the live page.
- Results (**PASS / FAIL / REVIEW / SKIPPED / ERROR**) are written back to a copy of your Excel file, with a confidence score, remarks and screenshots of failures.

---

## How it works

```
 Excel test case ──► Jev decides: action + element + value ──► Playwright runs it in Chrome
                                                                        │
 Excel result  ◄── PASS / FAIL / REVIEW ◄── Jev verifies the Expected Result on the page
```

For each step:

1. The runner collects the visible interactive elements on the page (links, buttons, inputs, dropdowns and so on).
2. **Jev decides** the action (`click`, `fill`, `select`, `press_enter` or `verify_only`), the target element and the test data value to use.
3. If Jev's confidence is below `DECISION_MIN_CONFIDENCE`, the step is **not executed**. It is marked **REVIEW** instead of being guessed.
4. Playwright performs the action.
5. If the step has an Expected Result, **Jev checks it** against the page text and form values and returns a probability:
   - `≥ 0.80` → **PASS**
   - `≤ 0.20` → **FAIL** (a screenshot is saved)
   - anything in between → **REVIEW**

If a step fails, errors or needs review before its action runs, the remaining steps of that test case are marked **SKIPPED**.

---

## Project structure

```
TypeSafe_Jev_Sample/
├── JevTestRunner.py          # the runner
├── manual_test_cases.xlsx    # sample test cases (saucedemo.com)
├── requirements.txt
├── .env.example              # copy to .env and add your API key
└── results/                  # created on run: result workbooks + screenshots
```

---

## Setup

**Prerequisites:** Python 3.10+, Google Chrome, and a TypeSafe API key.

```powershell
# 1. Create and activate a virtual environment
python -m venv .venv
.venv\Scripts\activate            # macOS/Linux: source .venv/bin/activate

# 2. Install dependencies
pip install -r requirements.txt
playwright install chromium       # only needed if you don't use installed Chrome

# 3. Add your API key
copy .env.example .env            # macOS/Linux: cp .env.example .env
# then open .env and set TYPESAFE_API_KEY=...
```

---

## Usage

```powershell
python JevTestRunner.py manual_test_cases.xlsx
```

With no file argument, the runner uses `manual_test_cases.xlsx`.

Output:

- `results/<name>_results_<timestamp>.xlsx`: your sheet with **Status**, **Confidence** and **Remarks** filled in and color-coded, plus a **Run Summary** sheet (counts per status and pass rate)
- `results/screenshots/`: full-page screenshots of every FAIL / ERROR step
- Console: a live step-by-step log, the number of Jev calls, input tokens and estimated cost

The workbook is saved after every test case, so a stopped run keeps the results it already has.

---

## Writing test cases

Put one row per **step**. Row 1 must have these headers (case-insensitive):

| Column | Required | Notes |
|---|---|---|
| TC ID | ✅ | Can be left blank on later steps; it is carried forward |
| Test Case Title | ✅ | Carried forward like TC ID |
| Start URL | ✅ | Opened in a fresh browser context at the start of each test case |
| Step No | ✅ | |
| Step Description | ✅ | Plain English, e.g. *"Click the Login button"* |
| Test Data | ✅ | A single value (`standard_user`) or several values (`username=standard_user; password=secret_sauce`) |
| Expected Result | ✅ | A statement about the page, e.g. *"Page URL contains inventory.html"*. Leave it blank to only perform the action |
| Status / Confidence / Remarks | ➖ | Added automatically if missing |

Example:

| TC ID | Test Case Title | Start URL | Step No | Step Description | Test Data | Expected Result |
|---|---|---|---|---|---|---|
| TC01 | Valid login | https://www.saucedemo.com/ | 1 | Enter the username | standard_user | Username field contains standard_user |
| | | | 2 | Enter the password | secret_sauce | |
| | | | 3 | Click the Login button | | Page URL contains inventory.html and the page lists products |

**Tip:** write each Expected Result as a statement about what is on screen, not about what the step did.

---

## Configuration

These settings are at the top of `JevTestRunner.py`:

| Setting | Default | Description |
|---|---|---|
| `HEADLESS` | `False` | `True` runs without a visible browser |
| `BROWSER_CHANNEL` | `"chrome"` | `"msedge"` for Edge, `None` for Playwright's bundled Chromium |
| `SLOW_MO_MS` | `0` | Delay between actions in ms (e.g. `400` for demos) |
| `DECISION_MIN_CONFIDENCE` | `0.60` | Below this confidence, the action is not executed and the step is marked REVIEW |
| `PASS_PROBABILITY` | `0.80` | Expected-result probability at or above this → PASS |
| `FAIL_PROBABILITY` | `0.20` | At or below this → FAIL |
| `MAX_ELEMENTS` | `80` | Interactive elements sent to Jev per step |
| `MAX_PAGE_TEXT` | `5000` | Characters of page text sent for verification |
| `SHEET_NAME` | `None` | Sheet to read (`None` = first sheet) |
| `COLUMNS` | — | Rename here if your Excel headers differ |

---

## Result statuses

| Status | Meaning |
|---|---|
| 🟢 **PASS** | Action done and the Expected Result was confirmed |
| 🔴 **FAIL** | The Expected Result was not met (a screenshot is saved) |
| 🟡 **REVIEW** | Low confidence or an unclear result; a human should check it |
| ⚪ **SKIPPED** | Not run because an earlier step in the test case did not pass |
| 🔴 **ERROR** | Exception or unreachable start URL (a screenshot is saved) |

---

## Security

- Keep your API key in `.env`. It is listed in `.gitignore`, so never commit it.
- Commit only `.env.example`, which holds a placeholder value.
