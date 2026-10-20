from behave import given, when, then
from pages.home_page import HomePage
from utils.excel_reader import load_test_data

def add_all_vegetables(context):
    for veg, qty in zip(context.vegetables, context.quantities):
        context.home.search(veg)
        context.home.set_quantity(qty)
        context.home.add_to_cart()
        context.home.clear_search()

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
    add_all_vegetables(context)

@given("the user has added vegetables to the cart")
def step_add_vegetables_to_cart(context):
    add_all_vegetables(context)

@then("the cart preview should contain the added vegetables")
def step_verify_preview(context):
    actual = context.home.get_cart_preview_items()
    actual_items = [
        (item["name"].split(" - ", 1)[0].strip().casefold(), item["quantity"])
        for item in actual
    ]
    expected_items = [
        (name.strip().casefold(), quantity)
        for name, quantity in zip(context.vegetables, context.quantities)
    ]
    assert actual_items == expected_items, (
        f"Expected cart items {expected_items}, found {actual_items}"
    )