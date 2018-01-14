using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using System;
using System.Linq;
using System.Windows;
using WpfApp_Selenium_KMU_20170904.Interface;

namespace WpfApp_Selenium_KMU_20170904.Abstract
{
    public abstract class AbstractCommonTools : ICommonTools
    {
        protected IWebDriver _drv;
        public AbstractCommonTools(IWebDriver drv)
        {
            _drv = drv;
            
        }
        abstract public void EnterToAccount();
        public IWebElement WaitUntilElementExist(string xpath, TimeSpan timeout)
        {
            try
            {
                WebDriverWait ww = new WebDriverWait(_drv, timeout);
                return ww.Until(ExpectedConditions.ElementExists(By.XPath(xpath)));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }

        }
        public void WaitUntilElementVisibleAndClick(string xpath)
        {
            try
            {
                IWebElement iwe = _drv.FindElement(By.XPath(xpath));
                if (iwe.Displayed == true)
                {
                    iwe.Click();
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }
        public bool FindAndWaitLast(string xPath, TimeSpan timeout)
        {
            bool res = false;
            try
            {
                WebDriverWait wdw = new WebDriverWait(_drv, timeout);
                res = wdw.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(xPath))).All(p => p.Displayed);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
            return res;
        }

    }
}
