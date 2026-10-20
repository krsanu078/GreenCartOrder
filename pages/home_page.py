import os
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException


class HomePage:
    URL = os.getenv("PAGE_URL", "https://rahulshettyacademy.com/seleniumPractise/")

    SEARCH_INPUT = (By.CSS_SELECTOR, "input.search-keyword")
    PRODUCT_CARDS = (By.CSS_SELECTOR, ".product")
    CART_ICON = (By.CSS_SELECTOR, 'a.cart-icon')
    CART_PRODUCT_NAMES = (By.XPATH, '//h4[@class="product-name"]')
    CART_PREVIEW_PRODUCT_NAMES = (By.CSS_SELECTOR, '.cart-preview .product-name')
    CART_PREVIEW_ITEMS = (By.CSS_SELECTOR, '.cart-preview li.cart-item')
    CART_PREVIEW_ITEM_NAME = (By.CSS_SELECTOR, "p.product-name")
    PROCEED_TO_CHECKOUT_BTN = (By.XPATH, '//button[text()="PROCEED TO CHECKOUT"]')

    def __init__(self, driver, timeout=10):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def open(self):
        self.driver.get(self.URL)

    def search(self, text):
        el = self.wait.until(EC.visibility_of_element_located(self.SEARCH_INPUT))
        el.clear()
        self.search_term = text
        el.send_keys(text)
        self.wait.until(lambda driver: self._find_product_card(driver, text) or False)

    def get_results_text(self):
        names = self.wait.until(
            lambda driver: self._read_texts(driver, self.CART_PRODUCT_NAMES) or False
        )
        return " ".join(names)

    @staticmethod
    def _read_texts(driver, locator):
        try:
            return [
                text
                for element in driver.find_elements(*locator)
                if (text := element.text.strip())
            ]
        except StaleElementReferenceException:
            return []

    def _find_product_card(self, driver, name):
        expected = name.strip().casefold()
        try:
            for card in driver.find_elements(*self.PRODUCT_CARDS):
                title = card.find_element(*self.CART_PRODUCT_NAMES).text
                product_name = title.split(" - ", 1)[0].strip().casefold()
                if product_name == expected and card.is_displayed():
                    return card
        except StaleElementReferenceException:
            return False
        return False

    def set_quantity(self, qty):
        for _ in range(3):
            try:
                card = self.wait.until(
                    lambda driver: self._find_product_card(driver, self.search_term)
                    or False
                )
                quantity = card.find_element(By.CSS_SELECTOR, "input.quantity")
                quantity.click()
                quantity.send_keys(Keys.CONTROL, "a")
                quantity.send_keys(Keys.BACKSPACE)
                quantity.send_keys(str(qty))
                self.wait.until(
                    lambda _: quantity.get_attribute("value") == str(qty)
                )
                return
            except StaleElementReferenceException:
                continue
        raise StaleElementReferenceException(
            f"Quantity element remained stale after 3 attempts: {self.search_term}"
        )

    def add_to_cart(self):
        card = self.wait.until(
            lambda driver: self._find_product_card(driver, self.search_term) or False
        )
        button = card.find_element(By.XPATH, './/button[normalize-space()="ADD TO CART"]')
        self.wait.until(lambda _: button.is_enabled())
        button.click()
        self.wait.until(lambda driver: self._product_was_added(driver, self.search_term))

    def _product_was_added(self, driver, name):
        try:
            card = self._find_product_card(driver, name)
            if not card:
                return False
            button = card.find_element(By.CSS_SELECTOR, ".product-action button")
            return "added" in button.get_attribute("class").casefold()
        except StaleElementReferenceException:
            return False

    def clear_search(self):
        el = self.wait.until(EC.visibility_of_element_located(self.SEARCH_INPUT))
        el.clear()

    def get_cart_preview_products(self):
        return [item["name"] for item in self.get_cart_preview_items()]

    def get_cart_preview_items(self):
        self.wait.until(EC.element_to_be_clickable(self.CART_ICON)).click()
        return self.wait.until(
            lambda driver: self._read_cart_preview_items(driver)
            or False
        )

    def _read_cart_preview_items(self, driver):
        try:
            items = []
            for item in driver.find_elements(*self.CART_PREVIEW_ITEMS):
                name = item.find_element(*self.CART_PREVIEW_ITEM_NAME).text.strip()
                quantity_text = item.find_element(
                    By.CSS_SELECTOR, "p.quantity"
                ).text.strip()
                if name and quantity_text:
                    items.append(
                        {
                            "name": name,
                            "quantity": int(quantity_text.split()[0]),
                        }
                    )
            return items
        except StaleElementReferenceException:
            return []

    def open_cart(self):
        self.wait.until(EC.element_to_be_clickable(self.CART_ICON)).click()

    def proceed_to_checkout(self):
        self.wait.until(EC.element_to_be_clickable(self.PROCEED_TO_CHECKOUT_BTN)).click()