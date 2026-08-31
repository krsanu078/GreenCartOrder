import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service as ChromeService
from selenium.webdriver.chrome.options import Options
from webdriver_manager.chrome import ChromeDriverManager

def before_all(context):
    options = Options()
    headless = os.getenv("HEADLESS", "1")
    if headless in ("1", "true", "True"):
        # Use Chrome headless mode
        # options.add_argument("--headless=new")
        options.add_argument("--disable-gpu")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
    options.add_argument("window-size=1920,1080")

    service = ChromeService(ChromeDriverManager().install())
    context.driver = webdriver.Chrome(service=service, options=options)

def after_all(context):
    try:
        context.driver.quit()
    except Exception:
        pass
