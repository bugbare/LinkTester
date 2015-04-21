using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

using NUnit.Framework;
using System.Net.Http;
using System.Configuration;
using System.Net.Http.Headers;
using System.Net;
using System.IO;

using Excel = Microsoft.Office.Interop.Excel;

namespace CrudXL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            xlWorkSheet.Cells[1, 1] = "Sheet 1 content";

            xlWorkBook.SaveAs("z:\\csharp-test-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file z:\\csharp-Excel.xls");
        }

        /* Read Excel Spreadsheet  */
        private void button2_Click(object sender, EventArgs e)
        {

            string crNumber = this.changeRequest.Text;
            string dcId = this.dataCentre.Text;

/*            if (sPattern == "") {
                sPattern = "http://www.cityindex.co.uk";
                MessageBox.Show("Site to test: " + sPattern);
            }
            else
            {
                MessageBox.Show("Site to test: " + sPattern);
            }*/

            
            DialogResult result = openFileDialog1.ShowDialog(); // Show the dialog.
	        if (result == DialogResult.OK) // Test result.
	        {
		        string file = openFileDialog1.FileName;       

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;
            

                HttpClientHandler httpClientHandler = new HttpClientHandler();
                httpClientHandler.AllowAutoRedirect = false;

                HttpResponseMessage response;

                string src;
                string exp;
                string act;
                int rCnt = 0;
                string resultFile = string.Format(crNumber+"-"+dcId+"-Redirects-{0:yyyy-MM-dd_hh-mm-ss-tt}.xls",
	                            DateTime.Now);

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                IWebDriver driver = new FirefoxDriver();

                range = xlWorkSheet.UsedRange;
                int tRows = range.Rows.Count;
                int tCols = range.Columns.Count;

                if (MessageBox.Show("TOTAL NUMBER OF CELLS TO QUERY: " + tRows*tCols + "\n" 
                    + "ROWS: " 
                    + tRows 
                    + " COLUMNS: " 
                    + tCols 
                    , "TOTAL ROWS Vs COLUMNS",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Asterisk,
                    MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                {
                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);
                    Application.Exit();
                    driver.Quit();
                }

                Excel.Workbook xlWorkBookNew;
                Excel.Worksheet xlWorkSheetNew;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBookNew = xlApp.Workbooks.Add(misValue);
                xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                xlWorkSheetNew.Cells[1, 1] = "SOURCE URL: ";
                xlWorkSheetNew.Cells[1, 2] = "EXPECTED DESTINATION URL";
                xlWorkSheetNew.Cells[1, 3] = "ACTUAL DESTINATION URL";
                xlWorkSheetNew.Cells[1, 4] = "HTTP RESPONSE CODE";
                xlWorkSheetNew.Cells[1, 5] = "RESULT";
           
                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {

                    src = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                    string sPattern = "http://";
                    

                        
                   if (System.Text.RegularExpressions.Regex.IsMatch(src, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    { 


                        HttpClient client = new HttpClient(httpClientHandler);
                        response = client.GetAsync(src).Result;
                        driver.Navigate().GoToUrl(src);
                        driver.Manage().Window.Maximize();
                        exp = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                        act = driver.Url.ToString();
                        try
                        {
                            Assert.AreEqual(HttpStatusCode.MovedPermanently, response.StatusCode);
                            StringAssert.AreEqualIgnoringCase(exp,act);
                            xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                            xlWorkSheetNew.Cells[rCnt, 1] = src;
                            xlWorkSheetNew.Cells[rCnt, 2] = exp;
                            xlWorkSheetNew.Cells[rCnt, 3] = act;
                            xlWorkSheetNew.Cells[rCnt, 4] = response.StatusCode.ToString();
                            xlWorkSheetNew.Cells[rCnt, 5] = "*****PASSED*****";
/*                            if (MessageBox.Show("SOURCE URL: "
                                + src
                                + "\nEXPECTED DESTINATION URL: "
                                + exp
                                + "\nACTUAL DESTINATION URL: "
                                + act 
                                +"\nROW: " 
                                + rCnt 
                                +"\nHTTP RESPONSE CODE: "
                                + response.StatusCode.ToString() 
                                + "\nLOCATION: "  
                                + response.Headers.Location.ToString() 
                                + "\nREASON: "
                                + response.ReasonPhrase.ToString()
                                + "\nCONNECTION: "
                                + response.Headers.Connection
                                , "*****PASSED*****",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.None,
                                MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                            {
                                xlWorkBookNew.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                                xlWorkBookNew.Close(true, misValue, misValue);
                                xlWorkBook.Close(true, null, null);
                                xlApp.Quit();
                                driver.Quit();

                                releaseObject(xlWorkSheet);
                                releaseObject(xlWorkBook);
                                releaseObject(xlWorkSheetNew);
                                releaseObject(xlWorkBookNew);
                                releaseObject(xlApp);

                                Application.Exit();
                            } */

                        }
                        catch (AssertionException AE)
                        {
                            xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                            xlWorkSheetNew.Cells[rCnt, 1] = src;
                            xlWorkSheetNew.Cells[rCnt, 2] = exp;
                            xlWorkSheetNew.Cells[rCnt, 3] = act;
                            xlWorkSheetNew.Cells[rCnt, 4] = response.StatusCode.ToString();
                            xlWorkSheetNew.Cells[rCnt, 5] = "*****FAILED*****";
                            if (MessageBox.Show("********************ERROR********************"
                                + "\nROW: "
                                + rCnt
                                + "\nSOURCE URL: "
                                + src
                                + "\nEXPECTED DESTINATION URL: "
                                + exp
                                + "\nACTUAL DESTINATION URL: "
                                + act  
                                + "\nHTTP RESPONSE CODE: "
                                + response.StatusCode.ToString() 
                                + "\nLOCATION: "  
                                + response.Headers.Location.ToString() 
                                + "\nREASON: "
                                + response.ReasonPhrase.ToString()
                                + "\nCONNECTION: "
                                + response.Headers.Connection
                                + "\n============================================="
                                + "\nNUNIT Says: "
                                + "\n"
                                + AE.ToString(),
                                "*****FAILED*****",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                            {
                                xlWorkBook.Close(true, null, null);
                                xlWorkBookNew.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                                xlWorkBookNew.Close(true, misValue, misValue);
                                xlApp.Quit();
                                driver.Quit();

                                releaseObject(xlWorkSheet);
                                releaseObject(xlWorkBook);
                                releaseObject(xlWorkSheetNew);
                                releaseObject(xlWorkBookNew);
                                releaseObject(xlApp);

                                Application.Exit();
                            }

/*                            if (MessageBox.Show("Do you want to quit?", "DEBUG", MessageBoxButtons.YesNo) == DialogResult.Yes)
                            {
                                xlWorkBookNew.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                                xlWorkBookNew.Close(true, misValue, misValue);
                                xlWorkBook.Close(true, null, null);
                                xlApp.Quit();
                                driver.Quit();

                                releaseObject(xlWorkSheet);
                                releaseObject(xlWorkBook);
                                releaseObject(xlWorkSheetNew);
                                releaseObject(xlWorkBookNew);
                                releaseObject(xlApp);

                                Application.Exit(); 
                            } */
                        }
                    }                    
                }


                MessageBox.Show("TESTS COMPLETED");
                xlWorkBook.Close(true, null, null);
                xlWorkBookNew.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookNew.Close(true, misValue, misValue);


                xlApp.Quit();
                driver.Quit();

                releaseObject(xlWorkSheetNew);
                releaseObject(xlWorkSheet);
                releaseObject(xlWorkBookNew);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);


                Application.Exit();
            }
        }


        private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                MessageBox.Show("Exception Occured while releasing object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void button3_Click_1(object sender, EventArgs e)
        {
            Application.Exit();
        }
    }
}
