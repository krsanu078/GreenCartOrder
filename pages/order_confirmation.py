from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

class OrderConfirmationPage:
    CONF_MSG = (By.XPATH, '//span[contains(text(),"Thank you, your order has been placed successfully")]')

    def __init__(self, driver, timeout=10):
        self.driver = driver
        self.wait = WebDriverWait(driver, timeout)

    def get_confirmation_message(self):
        el = self.wait.until(EC.visibility_of_element_located(self.CONF_MSG))
        return el.text