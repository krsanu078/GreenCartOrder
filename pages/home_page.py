import os
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class HomePage:
    URL = os.getenv("PAGE_URL", "https://rahulshettyacademy.com/seleniumPractise/")

    SEARCH_INPUT = (By.CSS_SELECTOR, "input.search-keyword")
    QUANTITY_INPUT = (By.XPATH, '//input[@class="quantity"]')
    ADD_TO_CART_BTN = (By.XPATH, '//button[text()="ADD TO CART"]')
    CART_ICON = (By.CSS_SELECTOR, "img[alt='Cart']")
    CART_PRODUCT_NAMES = (By.XPATH, '//p[@class="product-name"]')
    PROCEED_TO_CHECKOUT_BTN = (By.XPATH, "//button[text()='PROCEED TO CHECKOUT']")

    def __init__(self, driver, timeout=10):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def open(self):
        self.driver.get(self.URL)

    def search(self, text):
        el = self.wait.until(EC.visibility_of_element_located(self.SEARCH_INPUT))
        el.clear()
        el.send_keys(text)
        # wait for search results to appear
        self.wait.until(lambda d: len(d.find_elements(*self.CART_PRODUCT_NAMES)) > 0)

    def set_quantity(self, qty):
        q = self.wait.until(EC.visibility_of_element_located(self.QUANTITY_INPUT))
        q.clear()
        q.send_keys(str(qty))

    def add_to_cart(self):
        btn = self.wait.until(EC.element_to_be_clickable(self.ADD_TO_CART_BTN))
        btn.click()

    def clear_search(self):
        el = self.wait.until(EC.visibility_of_element_located(self.SEARCH_INPUT))
        el.clear()

    def get_cart_preview_products(self):
        # click cart icon then read names without proceeding
        self.driver.find_element(*self.CART_ICON).click()
        items = self.wait.until(EC.presence_of_all_elements_located(self.CART_PRODUCT_NAMES))
        names = [i.text for i in items]
        return names

    def open_cart(self):
        self.wait.until(EC.element_to_be_clickable(self.CART_ICON)).click()

    def proceed_to_checkout(self):
        self.wait.until(EC.element_to_be_clickable(self.PROCEED_TO_CHECKOUT_BTN)).click()