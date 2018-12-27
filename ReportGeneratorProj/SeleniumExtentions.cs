using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
namespace SeleniumExtentions
{
    public static class SeleniumExtentions
    {
        public static IWebElement UntilElementIsVisible(this WebDriverWait wait, string locator)
        {
            return wait.Until(c =>
            {
                var el = c.FindElement(By.CssSelector(locator));
                if (el.Displayed)
                {
                    return el;
                }
                else
                {
                    return null;
                }
            });
        }
    }
}
