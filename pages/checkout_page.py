from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.support.select import Select

class CheckoutPage:
    CHECKOUT_PRODUCT_NAMES = (By.CSS_SELECTOR, "p.product-name")
    PROMO_INPUT = (By.CLASS_NAME, "promoCode")
    PROMO_BTN = (By.CSS_SELECTOR, ".promoBtn")
    PROMO_INFO = (By.CSS_SELECTOR, "span.promoInfo")
    TOTAL_AMOUNT = (By.XPATH, '//span[@class="totAmt"]')
    ITEM_AMOUNT_CELLS = (By.XPATH, "//tr/td[5]/p")
    PLACE_ORDER_BTN = (By.XPATH, '//button[text()="Place Order"]')
    COUNTRY_SELECT = (By.CSS_SELECTOR, "select")
    PROCEED_BTN = (By.XPATH, '//button[text()="Proceed"]')
    TERMS_CHECKBOX = (By.XPATH, '//input[@type="checkbox"]')

    def __init__(self, driver, timeout=10):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def get_checkout_products(self):
        elems = self.wait.until(EC.presence_of_all_elements_located(self.CHECKOUT_PRODUCT_NAMES))
        return [e.text for e in elems]

    def apply_promo(self, code):
        promo_input = self.wait.until(EC.visibility_of_element_located(self.PROMO_INPUT))
        promo_input.clear()
        promo_input.send_keys(code)
        self.driver.find_element(*self.PROMO_BTN).click()
        # wait for promo info to update
        self.wait.until(EC.visibility_of_element_located(self.PROMO_INFO))

    def get_promo_message(self):
        return self.driver.find_element(*self.PROMO_INFO).text

    def get_total_amount(self):
        text = self.wait.until(EC.visibility_of_element_located(self.TOTAL_AMOUNT)).text
        return int(float(text))

    def get_item_amounts(self):
        elems = self.wait.until(EC.presence_of_all_elements_located(self.ITEM_AMOUNT_CELLS))
        return [int(float(e.text)) for e in elems]

    def place_order(self, country_value):
        # move to Place Order then click
        action = ActionChains(self.driver)
        place_btn = self.wait.until(EC.element_to_be_clickable(self.PLACE_ORDER_BTN))
        action.move_to_element(place_btn).click(place_btn).perform()
        # select country
        sel = Select(self.wait.until(EC.visibility_of_element_located(self.COUNTRY_SELECT)))
        sel.select_by_value(country_value)
        # accept terms
        self.wait.until(EC.element_to_be_clickable(self.TERMS_CHECKBOX)).click()
        # proceed
        self.wait.until(EC.element_to_be_clickable(self.PROCEED_BTN)).click()