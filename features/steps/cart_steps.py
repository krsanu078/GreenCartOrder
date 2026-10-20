from behave import when, then
from pages.checkout_page import CheckoutPage

@when("the user opens the cart and proceeds to checkout")
def step_open_cart_and_checkout(context):
    context.cart_preview_items = context.home.get_cart_preview_items()
    context.home.proceed_to_checkout()
    context.checkout = CheckoutPage(context.driver)

@then("the checkout items should match the cart preview items")
def step_compare_checkout_items(context):
    checkout_items = context.checkout.get_checkout_items()
    cart_items = [
        (item["name"].split(" - ", 1)[0].strip().casefold(), item["quantity"])
        for item in context.cart_preview_items
    ]
    checkout_pairs = [
        (item["name"].split(" - ", 1)[0].strip().casefold(), item["quantity"])
        for item in checkout_items
    ]
    assert cart_items == checkout_pairs, (
        f"Cart items {cart_items} != checkout items {checkout_pairs}"
    )
    for item in checkout_items:
        assert item["line_total"] == item["quantity"] * item["unit_price"], (
            f"Incorrect line total for {item['name']}: {item['line_total']}"
        )