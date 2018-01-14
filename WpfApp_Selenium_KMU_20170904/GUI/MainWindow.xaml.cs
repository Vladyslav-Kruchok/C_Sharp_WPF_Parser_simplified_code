using System;
using System.Collections.Generic;
using System.Windows;
using OpenQA.Selenium;
using OpenQA.Selenium.Support.UI;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Firefox;
using WpfApp_Selenium_KMU_20170904.MyOffice;
using ExcelMO = Microsoft.Office.Interop.Excel;
using WpfApp_Selenium_KMU_20170904.MyWeb;
using System.Collections.ObjectModel; //ReadOnlyCollection
using WpfApp_Selenium_KMU_20170904.MyWeb.Path;
using WpfApp_Selenium_KMU_20170904.MyWeb.Enums;
using WpfApp_Selenium_KMU_20170904.MyWeb.Store;
using System.Data;

namespace WpfApp_Selenium_KMU_20170904
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        IWebDriver Drv;
        int FinishRow;
        ExcelMO.Application appl;
        ExcelMO._Workbook xlWorkBook;
        UA UATools;
        NET OSS;

        Class_sssCollection Ssss = new Class_sssCollection();

        PassCollection pass = new PassCollection();

        PathContaintsTextCollection pctc = new PathContaintsTextCollection();

        MainMenuCollection mm = new MainMenuCollection();


        public MainWindow()
        {
            InitializeComponent();
            Init();
        }
        private void Init()
        {
            FinishRow = 0;
            appl = new ExcelMO.Application();
            UATools = new UA();
        }

        private void LoginUABT_Click(object sender, RoutedEventArgs e)
        {
            UATools.EnterLP(ref Drv);
        }

        private void StartUA_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.StartSeChromeUrl(MyUrl.UA, ref Drv);
            UATools.InitDrv(ref Drv);
        }

        private void button1_Click(object sender, RoutedEventArgs e)
        {
            Drv.Close();
            Drv.Quit();
        }

        private void button2_Click(object sender, RoutedEventArgs e)
        {
            UATools.MainMenu(ref Drv, mm[MainMenu.StEd], "Сп");
        }

        private void ScrollDown_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.ScrollDownListZdobuvachiv(ref FinishRow, ref Drv);
        }

        private void AccountBT_Click(object sender, RoutedEventArgs e)
        {
            int res = AddToolsSe.getXpathCount(xPathTB.Text, ref Drv);
            MessageBox.Show(res.ToString());
        }

        private void NextPageBT_Click(object sender, RoutedEventArgs e)
        {
            UATools.Next(ref Drv, 10);
        }
        private void PrevPageBT_Click(object sender, RoutedEventArgs e)
        {
            UATools.Prev(ref Drv);
        }

        private void FirstRowBT_Click(object sender, RoutedEventArgs e)
        {
            UATools.SubStrAttrTopFirstRow(ref Drv, "style", "top:", 4, "width:", 4);
        }
        private void LastBT_Click(object sender, RoutedEventArgs e)
        {
            UATools.SubStrAttrTopLastRow(ref Drv, "style", "top:", 4, "width:", 4);
        }

        private void checkXpathBT_Click(object sender, RoutedEventArgs e)
        {
            string xPath = xPathTB.Text;
            try
            {
                IWebElement First = Drv.FindElement(By.XPath(@xPath));
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }

        private void clickBT_Click(object sender, RoutedEventArgs e)
        {
            string xPath = xPathTB.Text;
            try
            {
                Drv.FindElement(By.XPath(@xPath)).Click();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());
            }
        }

        private void CloseUABT_Click(object sender, RoutedEventArgs e)
        {
            UATools.CloseUA(ref Drv);
        }

        private void GetListBT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                By xPathTitleOfTable = By.XPath(@"//div[@class='Xxx__Table__Row']/div");
                List<IWebElement> titleElements = AddToolsSe.getXpathElements(xPathTitleOfTable, ref Drv);
                ExcelMO._Worksheet workSheetFirst = MakeXLS.CreateExcelSingleSheet(ref appl);
                MakeXLS.MakeTitle(ref titleElements, ref workSheetFirst);

                string DataOfTable = @"//div[@aria-label='row'][1]/div";
                int countColumn = AddToolsSe.getXpathElements(By.XPath(DataOfTable), ref Drv).Count;
                MakeXLS.MakeData(ref Drv, DataOfTable, countColumn, ref workSheetFirst);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }

        private void SaveToXlsBT_Click(object sender, RoutedEventArgs e)
        {
            MakeXLS.SaveXls(@"D:\", "eWeb", ref appl);
        }


        private void StartEdDbBT_Click(object sender, RoutedEventArgs e)
        {
            UATools.StartSeChromeUrl(MyUrl.ED_UA, ref Drv);
            UATools.InitDrv(ref Drv);
        }

        private void SearchPI_Click(object sender, RoutedEventArgs e)
        {
            //1 searched
            UATools.SearchXPathAndClickEnterText(ref Drv, @"//input", PI_FSP.Text);
            PI_FSP.Clear();
            try
            {
                //2 click on checkbox
                UATools.SearchXPathAndClick(ref Drv, "//input[@type='checkbox']");
                //3 click on button 
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5));
                IWebElement ActToDo = Drv.FindElement(By.XPath("//div[@class='bottom-xs']/div/div[2]/div/div/button"));
                ActToDo.Click();
                //4 click on Doc
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5));
                IWebElement DocPI = Drv.FindElement(By.XPath("//span[@index]"));
                DocPI.Click();
                //5 ActToDo inside PI
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5));
                ReadOnlyCollection<IWebElement> buttons = Drv.FindElements(By.XPath("//button"));
                int size = buttons.Count;
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5));
                buttons[size - 1].Click();
                //6 Add Doc
                ReadOnlyCollection<IWebElement> span = Drv.FindElements(By.XPath("//span"));
                int size_span = span.Count;
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
                span[size_span - 1].Click();
            }
            catch (Exception)
            {
                ;
            }
        }

        private void PI_Click(object sender, RoutedEventArgs e)
        {
            UATools.MainMenu(ref Drv, mm[MainMenu.PI]);
        }

        private void EntBT_Click(object sender, RoutedEventArgs e)
        {
            UATools.MainMenu(ref Drv, mm[MainMenu.Ent]);
        }
        ReadOnlyCollection<IWebElement> RocXPath(ref IWebDriver Drv, string xPath, out int count)
        {
            ReadOnlyCollection<IWebElement> roc = Drv.FindElements(By.XPath(xPath));
            count = roc.Count;
            return roc;
        }
     
        private void SAVO_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UATools.AllElemContainText(ref Drv, "*", "Xxx", "ancestor::button").Click();
                UATools.AllElemContainText(ref Drv, "*", "Yyy", "following-sibling::div[2]/div").Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.2));
                UATools.AllElemContainText(ref Drv, "*", "Ccc", "ancestor::span").Click();
                UATools.AllElemContainText(ref Drv, "*", "Ddd", "following-sibling::input")
                    .SendKeys(Ssss[Class_sss.EAdm]);
                UATools.AllElemContainText(ref Drv, "*", "Eee", "following-sibling::input[@type='pass']")
                    .SendKeys(pass[Class_sss.EAdm]);

                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5));
                UATools.AllElemContainText(ref Drv, "*", "Fff", "ancestor::div[8]/div/div[2]/div/button").Click();
            }
            catch (Exception)
            {
                ;
            }

        }

        private void SR_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UATools.AllElemContainText(ref Drv, "*", "Eee", "ancestor::button").Click();

                UATools.AllElemContainText(ref Drv, "*", "Jjj", "following-sibling::div/div").Click();
                UATools.AllElemContainText(ref Drv, "*", "Ggg", "ancestor::span").Click();

                UATools.AllElemContainText(ref Drv, "*", "Hhh", "following-sibling::input").SendKeys(Ssss[Class_sss.ARr]);
                UATools.AllElemContainText(ref Drv, "*", "Kkk", "following-sibling::input[@type='pass']").SendKeys(pass[Class_sss.ARr]);

                UATools.AllElemContainText(ref Drv, "*", "Kkk", "following-sibling::div/input").SendKeys(Ssss[Class_sss.Sss]);
                UATools.AllElemContainText(ref Drv, "*", "Lll", "preceding::input[@type='pass'][1]").SendKeys(pass[Class_sss.Sss]);

                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(1));
                UATools.AllElemContainText(ref Drv, "*", "Zzz", "ancestor::div/div/div/button").Click();

            }
            catch (Exception)
            {
                ;
            }

        }



        private void ChO_BT_Click(object sender, RoutedEventArgs e)
        {
            Drv.Navigate().Refresh();
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.5));
            string numbOooo;
            try
            {
                WebDriverWait ww = new WebDriverWait(Drv, TimeSpan.FromSeconds(20));
                ww.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("//ul/li")));
                ReadOnlyCollection<IWebElement> colectionPage = Drv.FindElements(By.XPath("//ul/li"));
                int page = colectionPage.Count;
                
                for (int j = 0; j < (page - 2); j++)
                {
                    By xPathTitleOfTable = By.XPath("//div[@aria-label='row']");
                    List<IWebElement> titleElements = AddToolsSe.getXpathElements(xPathTitleOfTable, ref Drv);
                    int count_elem = 1;
                    int count = 0;
                    foreach (var item in titleElements)
                    {
                        int numbSymb = 10;
                        numbOooo = "";
                        
                        for (int i = 0; i < numbSymb; i++)
                        {
                            if (item.Text[i] == '/')
                            {
                                break;
                            }
                            numbOooo += item.Text[i];

                        }
                        if (numbOooo == NO_TB.Text)
                        {
                            RocXPath(ref Drv, "//input[@type='checkbox']", out count)[count_elem].Click();
                            UATools.AllElemContainText(ref Drv, "*", "АА", "following-sibling::div/button").Click();
                            UATools.AllElemContainText(ref Drv, "*", "КК", "ancestor::span").Click();
                            numbOooo = NO_TB.Text = "";
                            return;
                        }
                        count_elem++;
                    }
                    UATools.Next(ref Drv, 10);
                }
                

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
                numbOooo = NO_TB.Text = "";
            }
        }

        private void SortBy_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.SortBy().Click();

        }

        private void ChoiceTop_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.CheckBox(10, Table.TableRowTop);
        }

        private void ChoiceAll_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.CheckBox(10, Table.Header);
        }

        private void ChoiceTopAndCorrect_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.CheckBox(10, Table.TableRowTop);
            UATools.ActionToDoBT().Click();
        }

        private void ChoiceAllAndCorrect_BT_Click(object sender, RoutedEventArgs e)
        {
            if (DateI_PR.Text == "" | DateE_PR.Text == "")
            {
                MessageBox.Show("!!!");
                return;
            }
            //ЧЧЧ
            UATools.CheckBox(10, Table.Header);
            UATools.ActionToDoBT("Рд").Click();
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
            UATools.ChoiceField("Дв");
            UATools.InsertData(DateI_PR.Text, 10, ChoiceField.DataI);
            UATools.SaveMultyChoiceData();
            try
            {

                //ЧЧЧЧ
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                UATools.CheckBox(10, Table.TableRowTop);
                UATools.ActionToDoBT("Рд").Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                UATools.ChoiceField("Дд");
                UATools.InsertData(DateE_PR.Text, 10, ChoiceField.DataE);
                UATools.SaveMultyChoiceData();
                //ЧЧЧ
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                UATools.CheckBox(10, Table.TableRowTop);
                UATools.ActionToDoBT("Рд").Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                UATools.ChoiceField("Пк");
                UATools.InsertData("Вор", 10, ChoiceField.DiRrsP);
                UATools.SaveMultyChoiceData();
                //ЧЧЧ
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                UATools.CheckBox(10, Table.TableRowTop);
                UATools.ActionToDoBT("Рд").Click();
                System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
                UATools.ChoiceField("ППП");
                UATools.InsertData("ИИИ", 10, ChoiceField.DiRrsN);
                UATools.SaveMultyChoiceData();
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message); 
            }
            
        }

        private void Save_And_P_BT_Click(object sender, RoutedEventArgs e)
        {
            UATools.SaveMultyChoiceData();
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(0.3));
            ChoiceTop_BT_Click(sender, e);
            UATools.ActionToDoBT("ЧЧЧ").Click();
        }
        private void StartNET_BT_Click(object sender, RoutedEventArgs e)
        {
            OSS = new NET(new FirefoxDriver());
        }

        private void EnterToAcc_BT_Click(object sender, RoutedEventArgs e)
        {
            OSS.EnterToAccount();
        }


        private void ClickXPathOsv_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IWebElement iwe = OSS.DrvOsvita().FindElement(By.XPath(XPathOsv_TB.Text));
                if (iwe.Displayed)
                {
                    iwe.Click();
                }
                else
                {
                    MessageBox.Show("Display = " + iwe.Displayed.ToString());
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }

        private void SendKeysXPathOsv_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                OSS.DrvOsvita().FindElement(By.XPath(XPathOsv_TB.Text)).SendKeys(KeysOsv_TB.Text + Keys.Enter);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }

        }


        private void TakeCollectionOsv_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                CollectionOsv_TB.Text = OSS.TakeCollections(XPathOsv_TB.Text);
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        private void MoveToElemOsv_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                IWebElement iwe = OSS.DrvOsvita().FindElement(By.XPath(XPathOsv_TB.Text));
                Actions _event = new Actions(OSS.DrvOsvita());
                _event.MoveToElement(iwe).Click().Perform();
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.Message);
            }
        }

        public DataSet ds = new DataSet();
        public DataTable dt = new DataTable();

        private void OpenXlsFile_BT_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                string fullFnPath = "";
                if (ds==null)
                {
                    ds = new DataSet();
                }
                if (ds == null)
                {
                    dt = new DataTable();
                }
                MakeXLS.Open_Xls_File(ref ds, ref fullFnPath);
                dt = ds.Tables[0];
                XlsFilePath_TB.Text = fullFnPath;
            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.Message);
            }


        }

        private void Start_BT_Click(object sender, RoutedEventArgs e)
        {
            if (XlsFilePath_TB.Text == "")
            {
                MessageBox.Show("Не открыт файл с данными");
            }
            else
            {
                OSS.Proces(ref dt);
                XlsFilePath_TB.Text = "";
            }
        }

    }
}