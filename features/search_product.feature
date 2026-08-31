Feature: Search for a product
  As a customer
  I want to search for a vegetable and see matching results

  Scenario: Search for 'tomato'
    Given the user is on the home page
    When the user searches for "tomato"
    Then search results should contain "tomato"
