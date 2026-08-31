import os
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys

class HomePage:
    URL = os.getenv("PAGE_URL", "https://rahulshettyacademy.com/seleniumPractise/#/")

    def __init__(self, driver):
        self.driver = driver

    def open(self):
        self.driver.get(self.URL)

    def search(self, query):
        search_input = self.driver.find_element(By.CSS_SELECTOR, "input.search-keyword")
        search_input.clear()
        search_input.send_keys(query)
        search_input.send_keys(Keys.ENTER)

    def get_results_text(self):
        elems = self.driver.find_elements(By.CSS_SELECTOR, "h4.product-name")
        return " ".join([e.text for e in elems])
