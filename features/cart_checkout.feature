Feature: Cart and checkout verification
	As a shopper
	I want to verify that items in the cart preview match items on the checkout page

	Background:
		Given the test data is loaded from Excel
		And the user is on the home page
		And the user has added vegetables to the cart

	Scenario: Verify cart preview and checkout list match
		When the user opens the cart and proceeds to checkout
		Then the checkout items should match the cart preview items