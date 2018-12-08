Feature: PageDataClass
	Class to hold all of the data needed from each page

@PageDataStructures
Scenario: Add Accessibility issues to PageData class
	Given I need to save a new issue
	When I create new A11yPage class
	Then The new object should have all params filled
	
