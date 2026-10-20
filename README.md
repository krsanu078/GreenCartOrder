# GreenCart E2E Test Automation

A browser-based end-to-end test suite for the GreenKart grocery shopping demo. It uses Selenium WebDriver for browser interactions and Behave to express workflows as readable Gherkin features. The suite covers product search, cart operations, checkout, promo codes, and order confirmation.

## Highlights

- Behavior-driven scenarios in Gherkin, backed by reusable Python step definitions.
- Page Object Model to keep browser selectors and interactions out of feature files.
- Excel-driven product and quantity data.
- An ordered test runner for a complete shopper workflow.
- HTML reports with per-step status, duration, screenshots, logs, and failure tracebacks.

## Technology

| Component | Purpose |
| --- | --- |
| Python | Test implementation and runner |
| Behave | Gherkin feature execution and step matching |
| Selenium WebDriver | Chrome browser automation |
| webdriver-manager | ChromeDriver setup |
| openpyxl | Read test data from Excel |
| behave-html-formatter | Generate the HTML report and render attachments |

## Scenarios

The root-level `run_tests.py` runs the Excel-reader unit tests first, then invokes Behave once with these features in workflow order:

| Order | Feature | What it verifies |
| --- | --- | --- |
| 1 | `search_product.feature` | Search results contain the requested product. |
| 2 | `add_to_cart.feature` | Products and quantities from the workbook can be added and are visible in the cart. |
| 3 | `cart_checkout.feature` | Cart preview products match checkout products. |
| 4 | `promo_code.feature` | Valid, empty, and invalid promo codes produce the expected messages; item amounts sum to the displayed total. |
| 5 | `place_order.feature` | Checkout completes with the selected country and terms accepted, and an order confirmation is shown. |

## Project Structure

```text
features/                 Gherkin features, Behave hooks, and step definitions
pages/                    Selenium Page Objects
tests/                    Unit tests for Excel input validation
utils/excel_reader.py     Excel test-data reader
GreenCartOrder.xlsx       Sample product, quantity, promo, and country data
run_tests.py              Ordered full-suite runner
behave.ini                HTML formatter and report output configuration
requirements.txt          Python dependencies
.github/workflows/ci.yml  Starter GitHub Actions test workflow
```

`features/environment.py` creates a fresh WebDriver for each scenario, enables headless Chrome when `HEADLESS=1`, and closes the driver after the scenario. It also attaches a screenshot and step log to the HTML report after every step. Step definitions coordinate the feature workflow; page objects own the browser locators and interactions.

## Getting Started

### Prerequisites

- Python 3.10 or newer
- Google Chrome
- Internet access to install Python packages and access the demo site

`webdriver-manager` downloads a compatible ChromeDriver when the test run starts.

### Windows PowerShell

From the project root:

```powershell
python -m venv .venv
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python run_tests.py
```

If the virtual environment is already active, run `python run_tests.py` directly. The runner stops if unit tests fail and returns a nonzero exit code if either test stage fails.

### Run One Feature

```powershell
behave features\search_product.feature
```

This runs only the selected feature and replaces the same HTML report.

## HTML Report

Every run creates or replaces `behave-report.html` in the project root. Open that file in a browser to see feature and scenario results. Expand a step's attachments to inspect its screenshot and text log, including status and duration. Failed steps also include failure details and the Python traceback. The generated report is ignored by Git because it is a run artifact and can become large.

## Test Data

`GreenCartOrder.xlsx` is read from the project root. Its active worksheet uses these columns:

| Column | Meaning |
| --- | --- |
| `Item` | Product name to search for |
| `Quantities` | Quantity to add |
| `Promo_code` | Promo-code test data |
| `Country` | Checkout country value |

The workbook is included so the sample suite can run as-is. The reader checks the required headers and rejects missing, non-integer, or non-positive quantities instead of silently substituting a value.

## Configuration

The default target is the GreenKart demo site. Set `PAGE_URL` before running to use another compatible environment. For example, in PowerShell:

```powershell
$env:PAGE_URL = "https://example.test/"
python run_tests.py
```

The default data file is `GreenCartOrder.xlsx` in the project root. Override it with `TEST_DATA_FILE`:

```powershell
$env:TEST_DATA_FILE = "C:\test-data\GreenCartOrder.xlsx"
python run_tests.py
```

The target must expose compatible product search, cart, checkout, and promo-code controls because the page objects use site-specific selectors.

## CI Status

`.github/workflows/ci.yml` runs the same unit and ordered browser-test runner on pushes and pull requests, with `HEADLESS=1`. It uploads `behave-report.html` as a GitHub Actions artifact even if the tests fail; download the `behave-html-report` artifact from the workflow run summary. This is CI, not deployment/CD; deployment is intentionally outside this test project's current scope.

## Known Constraints

- Tests depend on an external demo site, so network availability and site changes can affect results.
- Browser selectors are specific to the current demo site.
- Screenshots are captured after every step, making the HTML report relatively large.
