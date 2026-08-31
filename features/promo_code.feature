Feature: Promo code validation and totals
	As a shopper
	I want to apply a promo code and validate the promo message and order totals

	Background:
		Given the test data is loaded from Excel
		And the user is on the home page
		And the user has added vegetables to the cart
		And the user proceeds to checkout

	Scenario Outline: Apply promo code and verify message & sums
		When the user applies the promo code "<code>"
		Then the promo message should be "<message>"
		And the sum of per-item amounts should equal the displayed total

		Examples:
			| code                 | message               |
			| rahulshettyacademy   | Code applied ..!      |
			|                       | Empty code ..!        |
			| invalidcode          | Invalid code ..!      |