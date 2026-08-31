from behave import when, then
from pages.checkout_page import CheckoutPage

@when('the user applies the promo code "{code}"')
def step_apply_promo(context, code):
    # make sure checkout page object exists (created earlier in flow)
    context.checkout = getattr(context, "checkout", CheckoutPage(context.driver))
    context.checkout.apply_promo(code)

@then('the promo message should be "{message}"')
def step_verify_promo_message(context, message):
    msg = context.checkout.get_promo_message()
    assert msg == message, f"Expected promo message '{message}', got '{msg}'"

@then("the sum of per-item amounts should equal the displayed total")
def step_verify_sum_equals_total(context):
    item_amounts = context.checkout.get_item_amounts()  # list of ints
    displayed_total = context.checkout.get_total_amount()
    assert sum(item_amounts) == displayed_total, f"Sum {sum(item_amounts)} != displayed total {displayed_total}"