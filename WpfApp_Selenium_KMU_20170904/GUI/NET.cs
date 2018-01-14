using OpenQA.Selenium;
using System;
using System.Windows;
using WpfApp_Selenium_KMU_20170904.Abstract;
using WpfApp_Selenium_KMU_20170904.CommonEnums;
using WpfApp_Selenium_KMU_20170904.CommonStore;
using WpfApp_Selenium_KMU_20170904.MyWeb;
using WpfApp_Selenium_KMU_20170904.MyWeb.NET.Enums;
using WpfApp_Selenium_KMU_20170904.MyWeb.NET.Path;
using System.Threading;
using ExcelMO = Microsoft.Office.Interop.Excel;
using System.Collections.ObjectModel;//ReadOnlyCollection
using System.Data;

namespace WpfApp_Selenium_KMU_20170904
{
    public class NET : AbstractCommonTools
    {
        private PathNETCollection _path = new PathNETCollection();
        public static ExcelMO.Application appl = new ExcelMO.Application();

        public NET(IWebDriver drv) : base(drv)
        {
            _drv.Navigate().GoToUrl(MyUrl.NET);
        }

        public override void EnterToAccount()
        {
            try
            {
                PassLoginCollection _passLogin = new PassLoginCollection();
                //1. click on the BT Вхід
                _drv.FindElement(By.XPath(_path[EnumPathNET.EBT]))
                    .Click();
                //2. send keys to input
                WaitUntilElementExist(_path[EnumPathNET.PITT], TimeSpan.FromSeconds(15))
                    .SendKeys(_passLogin[WebSite.Xxx_net] + Keys.Enter);
                Thread.Sleep(TimeSpan.FromSeconds(15));
                //3.choice university
                WaitUntilElementExist(_path[EnumPathNET.PHH], TimeSpan.FromSeconds(15))
                    .Click();
                Thread.Sleep(TimeSpan.FromSeconds(5));
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }
        public IWebDriver DrvOsvita()
        {
            return _drv;
        }
        public string TakeCollections(string xpath)
        {
            ReadOnlyCollection<IWebElement> RocIwe = _drv.FindElements(By.XPath(xpath));
            string res = "";
            int coount = 1;
            foreach (var item in RocIwe)
            {
                string im = item.Text;
                string newim = "";
                newim = im.Replace("\n", "");
                newim = newim.Replace("\r", "");
                if (im == "" || im == " ")
                {
                    res += $"[{coount.ToString()}] [-]\n";
                    coount++;
                    continue;
                }
                res += $"[{coount.ToString()}] {im}\n";
                coount++;
                //item.SendKeys("xxxxx");
            }
            return res;
        }
        private string CurentFSP_Anceta(int index)
        {
            string fName = "";
            string sName = "";
            string lName = "";
            string FSP = "";
            int count = 0;
            ReadOnlyCollection<IWebElement> RocIwe = _drv.FindElements(By.XPath($"//table[@id='Table']//child::tr[{index}]/td[2]/label"));

            foreach (var item in RocIwe)
            {
                count++;
                switch (count)
                {
                    case 1:
                        {
                            if ($"{item.Text.Trim()}" == "-")
                            {
                                fName = "";
                            }
                            else
                            {
                                fName = $"{item.Text.Trim()}";
                            }
                        }

                        break;
                    case 2:
                        {
                            if ($"{item.Text.Trim()}" == "-")
                            {
                                sName = "";
                            }
                            else
                            {
                                sName = $"{item.Text.Trim()}";
                            }
                        }

                        break;
                    case 3:
                        {
                            if ($"{item.Text.Trim()}" == "-")
                            {
                                lName = "";
                            }
                            else
                            {
                                lName = $"{item.Text.Trim()}";
                            }
                            FSP = ((fName + " " + sName).Trim() + " " + lName).Trim();
                        }
                        break;
                    default:
                        break;
                }
            }
            return FSP;
        }
        private int CountCollectionStAnketa()
        {
            
            ReadOnlyCollection<IWebElement> RocIwe;
            int count = 0;
            if (FindAndWaitLast(_path[EnumPathNET.LSAA], TimeSpan.FromSeconds(15)))
            {
                RocIwe = _drv.FindElements(By.XPath(_path[EnumPathNET.LSAA]));
                count = RocIwe.Count;
            }
            return count;
        }
        private string OnlyLetter(string str)
        {
            string res = "";
            foreach (var item in str)
            {
                switch (item)
                {
                    case '1':
                        continue;
                    case '2':
                        continue;
                    case '3':
                        continue;
                    case '4':
                        continue;
                    case '5':
                        continue;
                    case '6':
                        continue;
                    case '7':
                        continue;
                    case '8':
                        continue;
                    case '9':
                        continue;
                    case '0':
                        continue;
                    case '.':
                        continue;
                    default:
                        res += item;
                        break;
                }
            }
            return res;
        }

        private void Search(ref DataTable dt, ref string fullnameWEB, ref bool result)
        {
            int FSPColomnXls = 0;
            int maxRowXls = dt.Rows.Count;
            int maxColomnXls = dt.Columns.Count;
            ReadOnlyCollection<IWebElement> RocIwe = _drv.FindElements(By.XPath(_path[EnumPathNET.AOLA]));

            for (int indexRowXls = 1; indexRowXls < maxRowXls; indexRowXls++)
            {
                string FSP_XlsTable = OnlyLetter(dt.Rows[indexRowXls][FSPColomnXls].ToString());

                if (fullnameWEB == FSP_XlsTable)
                {
                    if ((RocIwe[(int)AAFF.Sity].GetAttribute("value") == "" && dt.Rows[indexRowXls][(int)XlsFileFields.Regg].ToString() == "")
                        | ((RocIwe[(int)AAFF.Sity].GetAttribute("value") == "-" && dt.Rows[indexRowXls][(int)XlsFileFields.Regg].ToString() == ""))
                        | ((RocIwe[(int)AAFF.Sity].GetAttribute("value") == "Київ" && dt.Rows[indexRowXls][(int)XlsFileFields.Regg].ToString() == "м.Київ")))
                    {
                        result = false;
                        break;
                    }
                    //заполнить необходимые поля на веб странице из данных таблицы
                    for (int colomn = 1; colomn < maxColomnXls; colomn++)
                    {
                        switch ((XlsFileFields)colomn)
                        {
                            case XlsFileFields.FSP:
                            case XlsFileFields.Diss:
                            case XlsFileFields.SityType:
                            case XlsFileFields.SttType:
                                continue;
                            case XlsFileFields.PCC:
                                RocIwe[(int)AAFF.PCC]
                                    .SendKeys(dt.Rows[indexRowXls][(int)XlsFileFields.PCC].ToString());
                                break;
                            case XlsFileFields.Regg:
                                {
                                    //click on Bt to choice a Regg
                                    _drv.FindElement(By.XPath(_path[EnumPathNET.TabAOLReggBT]))
                                        .Click();
                                    if (FindAndWaitLast(_path[EnumPathNET.TabAOLReggsList], TimeSpan.FromSeconds(15)))
                                    {
                                        ReadOnlyCollection<IWebElement> RocIweReggs = _drv.FindElements(By.XPath(_path[EnumPathNET.TabAOLReggsList]));
                                        //searched Regg
                                        string Regg = dt.Rows[indexRowXls][(int)XlsFileFields.Regg].ToString();
                                        //find Regg and click
                                        IWebElement reg = GetSearchItem(ref RocIweReggs, Regg);
                                        reg.Click();
                                    }

                                }
                                break;
                            case XlsFileFields.Sity:
                                {
                                    string sity = dt.Rows[indexRowXls][(int)XlsFileFields.Sity].ToString();
                                    string sityType = dt.Rows[indexRowXls][(int)XlsFileFields.SityType].ToString();
                                    if (sityType == "місто")
                                    {
                                        RocIwe[(int)AAFF.Sity].SendKeys(sity);
                                    }
                                    else
                                    {
                                        RocIwe[(int)AAFF.Sity].SendKeys($"{sityType} {sity}");
                                    }
                                }
                                break;
                            case XlsFileFields.Stt:
                                {
                                    string Stt = dt.Rows[indexRowXls][(int)XlsFileFields.Stt].ToString();
                                    string SttType = dt.Rows[indexRowXls][(int)XlsFileFields.SttType].ToString();
                                    if (SttType == "вул.")
                                    {
                                        RocIwe[(int)AAFF.Stt].SendKeys(Stt);
                                    }
                                    else
                                    {
                                        RocIwe[(int)AAFF.Stt].SendKeys($"{SttType} {Stt}");
                                    }
                                }
                                break;
                            default:
                                RocIwe[colomn - 3]
                                    .SendKeys(dt.Rows[indexRowXls][colomn].ToString());
                                break;
                        }
                    }
                    result = true;
                    return;
                }
            }
        }
        private IWebElement GetSearchItem(ref ReadOnlyCollection<IWebElement> RocIweReggs, string searchedItem)
        {
            IWebElement res = null;
            foreach (var item in RocIweReggs)
            {
                if (item.Text == searchedItem)
                {
                    return item;
                }
            }
            return res;
        }

        public void Proces(ref DataTable dt)
        {

            int count = CountCollectionStAnketa();
            // результат пошуку
            bool search_result = true;
            //Обрати  першого 
            WaitUntilElementExist(_path[EnumPathNET.TopCheckBox_Anketa], TimeSpan.FromSeconds(5))
                .Click();
            //Редагувати анкету
            WaitUntilElementExist(_path[EnumPathNET.EditSt_Anketa], TimeSpan.FromSeconds(5))
                .Click();
            Thread.Sleep(TimeSpan.FromSeconds(7));

            for (int i = 0; i < count; i++)
            {
                string fullNameWEB = CurentFSP_Anceta(i + 2);
                if (MessageBox.Show("Fill in????", fullNameWEB, MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {

                    WaitUntilElementVisibleAndClick(_path[EnumPathNET.AOLA]);
                    Thread.Sleep(TimeSpan.FromSeconds(0.5));
                    Search(ref dt, ref fullNameWEB, ref search_result);
                    if (!search_result)
                    {
                        //go to next Stt
                        WaitUntilElementExist(_path[EnumPathNET.NextStt_Anketa], TimeSpan.FromSeconds(15))
                            .Click();
                        Thread.Sleep(TimeSpan.FromSeconds(4));
                        continue;
                    }
                    WaitUntilElementVisibleAndClick(_path[EnumPathNET.PersonData]);

                    WaitUntilElementVisibleAndClick(_path[EnumPathNET.ChangePass_PersonData]);

                    //choice element
                    ReadOnlyCollection<IWebElement> RocPassVollections = _drv.FindElements(By.XPath(_path[EnumPathNET.CollOfDoc_PersonData]));
                    IWebElement reg = GetSearchItem(ref RocPassVollections, "PF");
                    reg.Click();
                    Thread.Sleep(TimeSpan.FromSeconds(0.5));
                }
                else
                {
                    //go to next Stt
                    WaitUntilElementExist(_path[EnumPathNET.NextStt_Anketa], TimeSpan.FromSeconds(15))
                        .Click();
                    Thread.Sleep(TimeSpan.FromSeconds(4));
                    continue;
                }

                if (MessageBox.Show("Save????", fullNameWEB, MessageBoxButton.YesNo, MessageBoxImage.Warning) == MessageBoxResult.Yes)
                {
                    //save Web
                    WaitUntilElementExist(_path[EnumPathNET.SaveDataStt_Anketa], TimeSpan.FromSeconds(15))
                        .Click();
                    //Thread.Sleep(TimeSpan.FromSeconds(10));

                }

                if (MessageBox.Show("Next Stt????", " ", MessageBoxButton.YesNo, MessageBoxImage.Question) == MessageBoxResult.Yes)
                {
                    //go to next Stt
                    WaitUntilElementExist(_path[EnumPathNET.NextStt_Anketa], TimeSpan.FromSeconds(15))
                        .Click();
                    //Thread.Sleep(TimeSpan.FromSeconds(5));
                }
            }
            MessageBox.Show("OK");
        }
       
    }
}
