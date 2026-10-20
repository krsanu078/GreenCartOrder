from behave import when, then

@when('the user searches for "{query}"')
def step_impl(context, query):
    context.home.search(query)

@then('search results should contain "{query}"')
def step_impl(context, query):
    results = context.home.get_results_text()
    assert query.lower() in results.lower()
