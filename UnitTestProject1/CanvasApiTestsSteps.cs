using System;
using TechTalk.SpecFlow;
using ReportGenerators;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [Binding]
    public class CanvasApiTestsSteps
    {
        private CourseInfo Course;
        private int id;
        [Given(@"I have entered course ID (.*)")]
        public void GivenIHaveEnteredCourseID(int p0)
        {
            id = p0;
        }

        [When(@"I create new CourseInfo object")]
        public void WhenICreateNewCourseInfoObject()
        {
            Course = new CourseInfo(id);

        }

        [Then(@"the new object should have all pages")]
        public void ThenTheNewObjectShowHaveAllPages()
        {

            Assert.AreEqual("Web Accessibility Compliance Resources", Course.CourseName);
            Assert.AreEqual(1026, Course.CourseIdOrPath);
            Assert.AreEqual("Web Accessibility Compliance Resources", Course.CourseCode);
            if (null == Course.PageHtmlList)
            {
                throw new Exception("PageList null");
            }
        }
    }
}
