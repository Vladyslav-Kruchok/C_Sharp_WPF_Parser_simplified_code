using OpenQA.Selenium;//IWebDriver
using OpenQA.Selenium.Support.UI;//WebDriverWait
using System;
using System.Collections.ObjectModel;//ReadOnlyCollection
using System.Linq;
using System.Windows;//MB
using System.Threading;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;
using WpfApp_Selenium_KMU_20170904.MyWeb.Path;
using WpfApp_Selenium_KMU_20170904.CommonStore;
using WpfApp_Selenium_KMU_20170904.CommonEnums;

namespace WpfApp_Selenium_KMU_20170904.MyWeb
{
    internal partial class UA
    {
        private IWebDriver _drv;
        PathContaintsTextCollection pctc = new PathContaintsTextCollection();
        SPSP rsp = new SPSP();
        LoginsCollection login = new LoginsCollection();
        PassLoginCollection passLogin = new PassLoginCollection();
        internal void InitDrv(ref IWebDriver Drv)
        {
            _drv = Drv;
        }
        internal IWebElement WaitUntil(int sec, string xPath)
        {
            WebDriverWait ww = new WebDriverWait(_drv, TimeSpan.FromSeconds(sec));
            return ww.Until(ExpectedConditions.ElementExists(By.XPath(xPath)));
        }
        internal void EnterLP(ref IWebDriver Drv)
        {
            try
            {
                IWebElement LogInput = Drv.FindElement(By.XPath(@"//input[@type='text']"));
                LogInput.SendKeys( login[WebSite.Xxx_ua]+ Keys.Enter);
                IWebElement PITT = Drv.FindElement(By.XPath(@"//input[@type='pass']"));
                PITT.SendKeys( passLogin[WebSite.Xxx_ua] + Keys.Enter);
                IWebElement Bt = Drv.FindElement(By.XPath(@"//button"));
                Bt.SendKeys(Keys.Enter);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        internal void StartSeChromeUrl(string url, ref IWebDriver Drv)
        {
            try
            {
                Drv = new OpenQA.Selenium.Chrome.ChromeDriver();
               // Drv.Manage().Window.Maximize();
                Drv.Navigate().GoToUrl(@url);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }

        }
        internal void MainMenu(ref IWebDriver Drv, string linkInMainMenu, string linkAfterInSide = "")
        {
            try
            {
                //1.find main button
                IWebElement MainMenu = Drv.FindElement(By.XPath(@"//button"));
                MainMenu.Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
                //2.find a link "З" 
                WebDriverWait ww = new WebDriverWait(Drv, TimeSpan.FromSeconds(15));
                IWebElement SttEdu = ww.Until(ExpectedConditions.ElementToBeClickable(By.LinkText(linkInMainMenu)));
                SttEdu.Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
                //3.find a link "Сп"
                if (linkAfterInSide != "")
                {
                    IWebElement edu3 = ww.Until(ExpectedConditions.ElementToBeClickable(By.LinkText(linkAfterInSide)));
                    edu3.Click();
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        internal void ScrollDownListZdobuvachiv(ref int FinishRow, ref IWebDriver Drv)
        {
            //scroll inside div ajax selenium
            try
            {
                if (FinishRow == 0)
                {
                    FinishRow = 870;
                }
                //atribute if row with data
                string firstPartText = "//div[@style='height: 0px; left: 0px; position: absolute; top: ";
                string lastPartText = "px; width: 100px; overflow: hidden; padding-right: 17px;']";
                string xPathRowTableBegin = String.Format(@"{0}{1}{2}", firstPartText, FinishRow, lastPartText);
                //heigth of row 30 px
                const int heightRow = 00;
                int nextFinishRow = ((AddToolsSe.getXpathCount(@"//div[@aria-label='row']", ref Drv) - 1) * heightRow) + FinishRow;
                if (nextFinishRow >= 0000)
                {
                    nextFinishRow = 0000;
                }
                //MessageBox.Show(getXpathCount("//div[@aria-label='row']").ToString());
                string xPathRowTableEnd = String.Format(@"{0}{1}{2}", firstPartText, nextFinishRow, lastPartText);

                AddToolsSe.ScrollToView(By.XPath(xPathRowTableBegin), ref Drv);
                WebDriverWait wait = new WebDriverWait(Drv, TimeSpan.FromSeconds(5));
                // wait for a new element at the end
                wait.Until((ExpectedConditions.ElementIsVisible(By.XPath(xPathRowTableEnd))));
                ///
                if (FinishRow < 0000)
                {
                    FinishRow = nextFinishRow;
                }
                else
                {
                    MessageBox.Show(FinishRow.ToString());
                }
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }
        internal void Prev(ref IWebDriver Drv)
        {
            try
            {
                bool isDisable = Drv.FindElement(By.XPath(@"//ul/li")).GetAttribute("class").Contains("dis");
                if (!isDisable)
                {
                    WebDriverWait ww = new WebDriverWait(Drv, TimeSpan.FromSeconds(10));
                    IWebElement prevPage = ww.Until(ExpectedConditions.ElementToBeClickable(By.XPath(@"//ul/li[1][@class=' mmm']")));
                    prevPage.Click();
                }
                else
                {
                    MessageBox.Show("First page");
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        internal void Next(ref IWebDriver Drv, int secWait)
        {
            try
            {
                WebDriverWait ww = new WebDriverWait(Drv, TimeSpan.FromSeconds(secWait));
                int count = AddToolsSe.getXpathCount(@"//ul/li", ref Drv);
                IWebElement nexPage = ww.Until(ExpectedConditions.ElementToBeClickable(By.XPath($"//ul/li[{count}]")));
                nexPage.Click();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        private string GetAttrSubStr(ref IWebElement First, string nameAttr, string firstPointAttr, int amountFirstSymb, string secondPointAttr, int amountSecondSymb)
        {
            int start = First.GetAttribute(nameAttr).ToString().IndexOf(firstPointAttr) + amountFirstSymb;
            int finish = First.GetAttribute(nameAttr).ToString().IndexOf(secondPointAttr) - amountSecondSymb;
            return First.GetAttribute(nameAttr).ToString().Substring(start, (finish - start));
        }
        internal string SubStrAttrTopFirstRow(ref IWebDriver Drv, string nameAttr, string firstPointAttr, int amountFirstSymb, string secondPointAttr, int amountSecondSymb)
        {
            try
            {
                IWebElement First = Drv.FindElement(By.XPath(@"//div[@class='Xxx__Table']/div/div/div[1]"));
                return GetAttrSubStr(ref First, nameAttr, firstPointAttr, amountFirstSymb, secondPointAttr, amountSecondSymb);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }
        internal string SubStrAttrTopLastRow(ref IWebDriver Drv, string nameAttr, string firstPointAttr, int amountFirstSymb, string secondPointAttr, int amountSecondSymb)
        {
            try
            {
                //account all last row
                int res = AddToolsSe.getXpathCount(@"//div[@aria-label='row']", ref Drv);
                string lastXpathRow = String.Format(@"//div[@class='Xxx__Table']/div/div/div[{0}]", res);
                IWebElement Last = Drv.FindElement(By.XPath(lastXpathRow));
                return GetAttrSubStr(ref Last, nameAttr, firstPointAttr, amountFirstSymb, secondPointAttr, amountSecondSymb);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                return null;
            }
        }
        internal void CloseUA(ref IWebDriver Drv)
        {
            string xPath = @String.Format("//*[@id=" + '"' + "content" + '"' + "]/div/div[2]/div/div/div[2]/button");
            try
            {
                Drv.FindElement(By.XPath(@xPath)).Click();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        internal void SearchXPathAndClickEnterText(ref IWebDriver Drv, string XPath, string textToInput = "" )
        {
            IWebElement we = Drv.FindElement(By.XPath(XPath));
            if (textToInput != "")
            {
                we.SendKeys(textToInput + Keys.Enter);
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
                try
                {
                    we.Clear();
                }
                catch (Exception)
                {
                    ;
                }

            }
        }
        internal void SearchXPathAndClick(ref IWebDriver Drv, string XPath)
        {
            WebDriverWait ww = new WebDriverWait(Drv, TimeSpan.FromSeconds(20));
            IWebElement we = ww.Until(ExpectedConditions.ElementExists(By.XPath(XPath)));
            we.SendKeys(XPath + Keys.Enter);
        }
        internal IWebElement AllElemContainText(ref IWebDriver Drv, string tag, params string[] searchedText)
        {
            int count = searchedText.Length;
            switch (count)
            {
                case 1:
                    {
                        string _XPath = String.Format($"//{tag}[contains(text(), '{searchedText[count - 1]}')]");
                        return WaitUntil(5, _XPath);
                    }
                case 2:
                    {
                        string _XPath = String.Format($"//{tag}[contains(text(), '{searchedText[count - 2]}')]//{searchedText[count - 1]}");
                        return WaitUntil(5, _XPath);
                    }
                default:
                    return Drv.FindElement(By.TagName("title"));
            }
        }
        internal IWebElement SortBy(string text = "П")
        {
            string xPathStr = String.Format($"//div[@aria-label='{text}']");
            IWebElement iwe = _drv.FindElement(By.XPath(xPathStr));
            return iwe;
        }
        internal void FindAndWaitLast(int sec, string xPath)
        {
            WebDriverWait wdw = new WebDriverWait(_drv, TimeSpan.FromSeconds(sec));
            wdw.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(xPath)));
        }
        internal bool FindAndWaitLast(int sec, string xPath, bool res)
        {
            WebDriverWait wdw = new WebDriverWait(_drv, TimeSpan.FromSeconds(sec));
            res = wdw.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath(xPath))).All(p=> p.Displayed);
            return res;
        }

        internal void FindAndElementExist(int sec, string xPath)
        {
            WebDriverWait wdw = new WebDriverWait(_drv, TimeSpan.FromSeconds(sec));
            wdw.Until(ExpectedConditions.ElementExists(By.XPath(xPath)));
        }
        internal void ClickOnSingleElement(string xPath)
        {
            _drv.FindElement(By.XPath(xPath)).Click();
        }
        internal void CheckBox(int sec, Table row)
        {
            try
            {
                string xPatnStr = "\0";
                ReadOnlyCollection<IWebElement> RocIwe;
                IWebElement iwe;
                switch (row)
                {
                    case Table.Header:
                        xPatnStr = String.Format("//div[@role='header']//child::input");
                        iwe = FindAndWaitLast(10, "//div[contains(@class, 'Container')]", false) ? _drv.FindElement(By.XPath(xPatnStr)) : null;
                        Thread.Sleep(TimeSpan.FromSeconds(0.5));
                        iwe.Click();
                        break;
                    case Table.TableRowTop:
                        xPatnStr = String.Format("//div[@class='Xxx_Grid_Container']//child::input");
                        RocIwe = FindAndWaitLast(10, "//div[contains(@class, 'Container')]", false) ? _drv.FindElements(By.XPath(xPatnStr)) : null;
                        Thread.Sleep(TimeSpan.FromSeconds(0.5));
                        RocIwe[0].Click();
                        break;
                    default:
                        break;
                }
            }
            catch (Exception)
            {

                throw;
            }

        }
        internal IWebElement Input(int sec, ref ChoiceField field)
        {
            try
            {
                return AllElemContainText(ref _drv, "*", "Р", rsp[field]);
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return null;
            }

        }
        internal IWebElement ActionToDoBT(string subMenuItem = "Р")
        {
            try
            {
                IWebElement mainBT = _drv.FindElement(By.XPath("//*[contains(text(),'А')] //following-sibling::div/button"));
                mainBT.Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                string xPathSubBT = String.Format($"//*[contains(text(),'{subMenuItem}')] //ancestor::span");
                FindAndElementExist(10, xPathSubBT);
                IWebElement subMenu = _drv.FindElement(By.XPath(xPathSubBT));
                return subMenu;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
                return null;
            }


        }
        internal void ChoiceField(string field)
        {
            string xPathChoiceDiv = String.Format("//*[contains(text(),'П')] //following-sibling::div/div");
            string xPathChoiceFealdToEdit = String.Format($"//*[contains(text(),'{field}')]//ancestor::span");
            FindAndElementExist(10, xPathChoiceDiv);
            ClickOnSingleElement(xPathChoiceDiv);
            FindAndElementExist(10, xPathChoiceFealdToEdit);
            ClickOnSingleElement(xPathChoiceFealdToEdit);
        }
        internal void InsertData(string data, int sec, ChoiceField field)
        {
            IWebElement input = Input(sec, ref field);
            input.Click();
            input.SendKeys(Keys.Home);
            input.SendKeys(data + Keys.Enter);   
        }
        internal void SaveMultyChoiceData()
        {
            try
            {
                IWebElement iwe = _drv.FindElement(By.XPath("//*[contains(text(),'З')] //ancestor::button"));
                iwe.Click();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


    }
}
