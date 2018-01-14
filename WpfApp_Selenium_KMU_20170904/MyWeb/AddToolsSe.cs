using OpenQA.Selenium;//Se
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows;//MB

namespace WpfApp_Selenium_KMU_20170904.MyWeb
{
    internal class AddToolsSe
    {
        internal static IWebElement ScrollToView(string xPath, ref IWebDriver Drv)
        {
            IWebElement elem;
            IJavaScriptExecutor jse;
            By lCh;
            try
            {
                lCh = By.XPath(xPath);
                elem = Drv.FindElement(lCh);
                jse = (IJavaScriptExecutor)Drv;
                jse.ExecuteScript("arguments[0].scrollIntoView(true);", elem);
                return elem;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return null;
            }
        }
        internal static void ScrollToView(By xPath, ref IWebDriver Drv)
        {
            IWebElement elem;
            IJavaScriptExecutor jse;
            try
            {
                elem = Drv.FindElement(xPath);
                jse = (IJavaScriptExecutor)Drv;
                jse.ExecuteScript("arguments[0].scrollIntoView(true);", elem);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
                return;
            }
        }
        internal static List<IWebElement> getXpathElements(string element, ref IWebDriver Drv)
        {
            List<IWebElement> located_element = Drv.FindElements(By.XPath(element)).ToList();
            return located_element;
        }
        internal static List<IWebElement> getXpathElements(By elementBy, ref IWebDriver Drv)
        {
            List<IWebElement> located_element = Drv.FindElements(elementBy).ToList();
            return located_element;
        }
        internal static int getXpathCount(string xPathElem, ref IWebDriver Drv)
        {
            int count = 0;
            List<IWebElement> located_element = Drv.FindElements(By.XPath(xPathElem)).ToList();
            foreach (var item in located_element)
            {
                count++;
            }
            return count;
        }
    }
    
}
