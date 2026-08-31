# Selenium BDD scaffold (behave)

This branch adds a basic Selenium + behave project scaffold.

Quick start

1. Create a virtual environment and install dependencies:

   python -m venv .venv
   source .venv/bin/activate   # or .venv\\Scripts\\activate on Windows
   pip install -r requirements.txt

2. Point the tests to your app (optional):

   - Edit pages/home_page.py and set PAGE_URL, or set the PAGE_URL environment variable.

3. Run the tests locally (headed):

   export HEADLESS=0
   behave

CI

- The included GitHub Actions workflow runs behave in headless Chrome.

Notes & next steps

- Replace the placeholder PAGE_URL with your application URL or set PAGE_URL env var.
- Add more feature files under features/ and implement page objects inside pages/.
- Improve waits (use WebDriverWait) and test data handling.
