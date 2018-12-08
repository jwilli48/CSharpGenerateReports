Feature: CanvasApiTests

Scenario: Create a CourseInfo object from Canvas course ID
	Given I have entered course ID 1026
	When I create new CourseInfo object
	Then the new object should have all pages