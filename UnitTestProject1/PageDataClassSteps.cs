using System;
using TechTalk.SpecFlow;
using ReportGenerators;
using Microsoft.VisualStudio.TestTools.UnitTesting;

namespace UnitTestProject1
{
    [Binding]
    public class PageDataClassSteps
    {
        private PageData A11yPage;
        [Given(@"I need to save a new issue")]
        public void GivenINeedToSaveANewIssue()
        {
            return;
        }

        [When(@"I create new A11yPage class")]
        public void WhenICreateClass()
        {
            A11yPage = new PageA11yData("test", "link", "", "asdf", "Empty href", 1);
        }
        
        [Then(@"The new object should have all params filled")]
        public void ThenTheNewObjectShouldHaveAllParamsFilled()
        {
            Assert.AreEqual("test", A11yPage.Location);
            Assert.AreEqual("link", A11yPage.Element);
            Assert.AreEqual("", A11yPage.Id);
            Assert.AreEqual("asdf", A11yPage.Text);
            Assert.AreEqual("Empty href", (A11yPage as PageA11yData).Issue);
            Assert.AreEqual(1, (A11yPage as PageA11yData).Severity);
        }
    }
}
