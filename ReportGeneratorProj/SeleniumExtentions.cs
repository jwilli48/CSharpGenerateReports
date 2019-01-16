using System.Linq;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
namespace SeleniumExtentions
{
    public static class SeleniumExtentions
    {   //Class to contain any extenstions to selenium
        public static IWebElement UntilElementIsVisible(this WebDriverWait wait, By locator)
        {   //Replace constantly used code that I was writing
            //Waits until the specified element is displayed or times out
            return wait.Until(c =>
            {
                try
                {
                    var el = c.FindElement(locator);
                    if (el.Displayed)
                    {
                        return el;
                    }
                    else
                    {
                        return null;
                    }
                }
                catch
                {
                    return null;
                }
            });
        }

        public static IWebElement UntilElementExist(this WebDriverWait wait, By locator)
        {
            return wait.Until(c =>
            {
                try
                {
                    var el = c.FindElement(locator);
                    return el;
                }
                catch
                {
                    return null;
                }
            });
        }

        public static IWebElement ReturnClick(this IWebElement element)
        {
            element.Click();
            return element;
        }

        public static IWebElement ForChildElement(this WebDriverWait wait, IWebElement element, By locator)
        {
            return wait.Until(c =>
            {
                try
                {
                    var el = element.FindElement(locator);
                    return el;
                }
                catch
                {
                    return null;
                }
            });
        }
    }
}
