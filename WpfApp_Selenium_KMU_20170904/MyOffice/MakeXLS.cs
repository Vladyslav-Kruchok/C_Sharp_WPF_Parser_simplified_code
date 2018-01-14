using OpenQA.Selenium;//IWebElement
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using ExcelMO = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using WpfApp_Selenium_KMU_20170904.MyWeb;
using System.Windows;
using Microsoft.Win32; //OpenFileDialog
using System.IO;//FileStream
using System.Data;
/*
It is important to note that every reference to an Excel COM object
had to be set to null when you have finished with it, including Cells, Sheets, everything.
The Marshal class is in the System.Runtime.InteropServices namespace,
so you should import the following namespace.
*/

namespace WpfApp_Selenium_KMU_20170904.MyOffice
{
    internal static class MakeXLS
    {

        private static object misValue = System.Reflection.Missing.Value;
        internal static ExcelMO._Worksheet CreateExcelSingleSheet(ref ExcelMO.Application appl)
        {
            // Make the object visible.
            appl.Visible = true;

            // Create a new, empty workbook and add it to the collection returned 
            // by property Workbooks. The new workbook becomes the active workbook.
            // Add has an optional parameter for specifying a praticular template. 
            // Because no argument is sent in this example, Add creates a new workbook. 
            appl.Workbooks.Add();

            // This example uses a single workSheet. The explicit type casting is
            // removed in a later procedure.
            return (ExcelMO.Worksheet)appl.ActiveSheet;
        }
        private static string FillRow52(int figureLetter = 65)
        {
            string collumnLetter = "";
            if (figureLetter > 64 & figureLetter < 91)
            {
                collumnLetter = String.Format("{0}", (char)figureLetter);
            }
            if (figureLetter > 90 & figureLetter < 117)
            {
                collumnLetter = String.Format("A{0}", (char)(figureLetter - 26));
            }
            return collumnLetter;
        }
        internal static void MakeTitle(ref List<IWebElement> titleList, ref ExcelMO._Worksheet workSheet)
        {
            int rowNumber = 1;
            int bigLetter = 65;//A
            foreach (var item in titleList)
            {
                workSheet.Cells[rowNumber, FillRow52(bigLetter)] = item.Text;
                bigLetter++;
            }
            return;
        }
        internal static void MakeData(ref IWebDriver Drv, string xPathDataTable, int countColumnDataTable, ref ExcelMO._Worksheet workSheet)
        {
            int rowNumber = 2;
            int bigLetter = 65;//A
            for (int i = 1; i < countColumnDataTable; i++)
            {
                string path = String.Format("{0}[{1}]", xPathDataTable, i);
                IWebElement el = AddToolsSe.ScrollToView(path, ref Drv);
                if (String.IsNullOrWhiteSpace(el.Text))
                {
                    workSheet.Cells[rowNumber, FillRow52(bigLetter)] = el.FindElement(By.TagName("path")).GetAttribute("d").Substring(0, 3) + "\n";
                    bigLetter++;
                }
                else
                {
                    workSheet.Cells[rowNumber, FillRow52(bigLetter)] = el.Text;
                    bigLetter++;
                }
            }
            return;
        }
        internal static void SaveXls(string path, string fileName, ref ExcelMO.Application appl)
        {
            try
            {
                Object fullPath = path + @"\" + fileName;
                appl.ActiveWorkbook.SaveAs(fullPath);
                //close application
                appl.ActiveWorkbook.Close();
                appl.Quit();
                Marshal.ReleaseComObject(appl);
                return;
            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());
            }
        }
        internal static bool? OpenXlsFile(ref ExcelMO.Application appl, ref string filename, ref ExcelMO._Workbook xlWorkBook)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";
            bool? resofd = ofd.ShowDialog();
            filename = ofd.FileName;
            if (resofd == true)
            {
                int fileformat = ofd.SafeFileName.IndexOf(".xlsx");
                try
                {
                    if (fileformat > -1)
                    {
                        //2007 format *.xlsx
                        xlWorkBook = appl.Workbooks.OpenXML(ofd.FileName,
                            Type.Missing);
                    }
                    else
                    {
                        //97-2003 format *.xls
                        xlWorkBook = appl.Workbooks.Open(ofd.FileName,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing,
                            Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                    }
                }
                catch (Exception ex)
                {

                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                ofd = null;
            }
            return resofd;
        }
        internal static bool? Open_Xls_File(ref DataSet ds, ref string fn)
        {
            OpenFileDialog ofd = new OpenFileDialog();
            ofd.DefaultExt = "*.xls;*.xlsx";
            ofd.Filter = "Excel 2003(*.xls)|*.xls|Excel 2007(*.xlsx)|*.xlsx";
            ofd.Title = "Выберите документ для загрузки данных";
            bool? resofd = ofd.ShowDialog();
            if (resofd == true)
            {
                int fileformat = ofd.SafeFileName.IndexOf(".xlsx");
                FileStream stream = null;
                Excel.IExcelDataReader IEDR;
                try
                {
                    stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read);
                    //Show path in texbox
                    fn = ofd.FileName.ToString();
                    if (fileformat > -1)
                    {
                        //2007 format *.xlsx
                        //IEDR = Excel.ExcelReaderFactory.CreateOpenXmlReader(stream);
                        IEDR = null;
                        MessageBox.Show("Not Implemented");
                    }
                    else
                    {
                        //97-2003 format *.xls
                       IEDR = Excel.ExcelReaderFactory.CreateBinaryReader(stream);
                    }
                    //Если данное значение установлено в true
                    //то первая строка используется в качестве 
                    //заголовков для колонок
                    IEDR.IsFirstRowAsColumnNames = false;
                    ds = IEDR.AsDataSet();
                    
                }
                catch (Exception ex)
                {
                    stream = null;
                    ofd = null;
                    MessageBox.Show(ex.ToString());
                }
            }
            else
            {
                ofd = null;
            }
            return resofd;
        }
    }
}
