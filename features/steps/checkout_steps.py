from behave import when, then
from pages.checkout_page import CheckoutPage
from pages.order_confirmation_page import OrderConfirmationPage

@when('the user places the order using the country from the data file')
def step_place_order(context):
    context.checkout = getattr(context, "checkout", CheckoutPage(context.driver))
    country = context.test_data.get("country", "")
    context.checkout.place_order(country)

@then('the confirmation should contain "placed successfully"')
def step_verify_order_confirmation(context):
    confirmation = OrderConfirmationPage(context.driver)
    msg = confirmation.get_confirmation_message()
    assert "placed successfully" in msg.lower()