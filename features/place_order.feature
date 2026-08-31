Feature: Place order
	As a shopper
	I want to place the order choosing a country and accepting terms

	Background:
		Given the test data is loaded from Excel
		And the user is on the home page
		And the user has added vegetables to the cart
		And the user proceeds to checkout

	Scenario: Complete checkout and verify confirmation
		When the user places the order using the country from the data file
		Then the confirmation should contain "placed successfully"