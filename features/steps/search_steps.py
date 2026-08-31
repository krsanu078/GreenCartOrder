from behave import given, when, then
from pages.home_page import HomePage
import time

@given('the user is on the home page')
def step_impl(context):
    context.home = HomePage(context.driver)
    context.home.open()

@when('the user searches for "{query}"')
def step_impl(context, query):
    context.home.search(query)
    time.sleep(1)

@then('search results should contain "{query}"')
def step_impl(context, query):
    results = context.home.get_results_text()
    assert query.lower() in results.lower()
