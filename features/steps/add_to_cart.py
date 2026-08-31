from behave import given, when, then
from pages.home_page import HomePage
from utils.excel_reader import load_test_data

@given("the test data is loaded from Excel")
def step_load_data(context):
    context.test_data = load_test_data()  # dict with vegetables, quantities, promocode, country
    context.vegetables = context.test_data["vegetables"]
    context.quantities = context.test_data["quantities"]
    context.promocode = context.test_data.get("promocode", "")
    context.country = context.test_data.get("country", "")

@given("the user is on the home page")
def step_open_home(context):
    context.home = HomePage(context.driver)
    context.home.open()

@when("the user searches and adds all vegetables from the data file")
def step_add_all(context):
    for veg, qty in zip(context.vegetables, context.quantities):
        context.home.search(veg)
        context.home.set_quantity(qty)
        context.home.add_to_cart()
        context.home.clear_search()

@then("the cart preview should contain the added vegetables")
def step_verify_preview(context):
    preview = context.home.get_cart_preview_products()
    # Normalize names for comparison
    preview_names = [p.lower().strip() for p in preview]
    expected = [v.lower().strip() for v in context.vegetables]
    for exp in expected:
        assert any(exp in p for p in preview_names), f"{exp} not found in cart preview: {preview_names}"