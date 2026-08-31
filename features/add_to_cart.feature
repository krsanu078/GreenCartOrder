Feature: Add vegetables to cart
	As a shopper
	I want to search for vegetables and add them to the cart with requested quantities

	Background:
		Given the test data is loaded from Excel
		And the user is on the home page

	Scenario: Add multiple vegetables from the test data
		When the user searches and adds all vegetables from the data file
		Then the cart preview should contain the added vegetables