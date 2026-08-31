from behave import when, then
from pages.home_page import HomePage
from pages.checkout_page import CheckoutPage

@when("the user opens the cart and proceeds to checkout")
def step_open_cart_and_checkout(context):
    context.home.open_cart()
    context.cart_preview_items = context.home.get_cart_preview_products()
    context.home.proceed_to_checkout()
    context.checkout = CheckoutPage(context.driver)

@then("the checkout items should match the cart preview items")
def step_compare_checkout_items(context):
    checkout_items = context.checkout.get_checkout_products()
    # Normalize text
    cp = [p.lower().strip() for p in context.cart_preview_items]
    co = [p.lower().strip() for p in checkout_items]
    assert cp == co, f"Cart preview {cp} != checkout list {co}"