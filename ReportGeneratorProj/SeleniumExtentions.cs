﻿using System;
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
    {   //Class to contain any extenstions to selenium
        public static IWebElement UntilElementIsVisible(this WebDriverWait wait, string locator)
        {   //Replace constantly used code that I was writing
            //Waits until the specified element is displayed or times out
            return wait.Until(c =>
            {
                try
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
                }
                catch
                {
                    return null;
                }
            });
        }
    }
}
