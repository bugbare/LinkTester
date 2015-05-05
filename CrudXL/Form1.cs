using NUnit.Framework;

using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Tokenizer;

using Excel = Microsoft.Office.Interop.Excel;

namespace CrudXL
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void createButton_Click(object sender, EventArgs e)
        {
            /* Instantiate Excel App object to use to generate an excell workbook*/
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            String pageUrl;

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }


           if (this.inputPageUrl.Text == "") 
                {
                    pageUrl = "http://www.cityindex.co.uk";
                }
            else 
                {
                    pageUrl = this.inputPageUrl.Text;
                }

            /* Instantiate all excel app objects to contain all the excel workbook components*/
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            /* a Value type missing object is required when generating a Workbook using the Excel.Workbook.SaveAs serializing method
             * to ignore certain options that are not required to be used - null is not allowed to be used instead */
            object misValue = System.Reflection.Missing.Value;
            /* Instantiate a workbook and initialise the first worksheet, assigned to the xlWorksheet object*/
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
            /* Initialise a new webdriver driver object and go to the initial landing page*/
            IWebDriver driver = new FirefoxDriver();
            driver.Navigate().GoToUrl(pageUrl);
            string pageTitle = driver.Title;
            string resultFile = string.Format("Test Data-{0:yyyy-MM-dd_hh-mm-ss-tt}.xls",
                                DateTime.Now);
            /* Create a local variable to hold contents of the page source*/

            var pageLinks = new List<string>();
            var anchors = driver.FindElements(By.TagName("a"));
            int rCnt = 1;
            string lPattern = "http";
            foreach (var element in anchors)
            {
                if (element.GetAttribute("href") != null)
                {
                    string pageLink = element.GetAttribute("href");
         
                    
                    if (System.Text.RegularExpressions.Regex.IsMatch(pageLink, lPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {
                        if (System.Text.RegularExpressions.Regex.IsMatch(pageLink, "/#", System.Text.RegularExpressions.RegexOptions.IgnoreCase) == false)
                        {
                            pageLinks.Add(pageLink);
                            xlWorkSheet.Cells[rCnt, 1] = pageLink;
                            rCnt++;
                        }
                    }
                }
            }

            rCnt = 1;
            foreach (var link in pageLinks)
            {
                driver.Navigate().GoToUrl(link);
                var destinationUrl = driver.Url;
                xlWorkSheet.Cells[rCnt, 2] = destinationUrl;
                driver.Navigate().Back();
                rCnt++;
            }



            xlWorkBook.SaveAs("z:\\" + resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();
            driver.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file at z:\\"+resultFile);
        }

        /* Read Excel Spreadsheet  */
        private void readButton_Click(object sender, EventArgs e)
        {

            string crNumber;
            string dcId;

            if (this.changeRequest.Text == null) 
                {
                    crNumber = "";
                }
            else 
                {
                    crNumber = this.changeRequest.Text;
                }
             
            if (this.dataCentre.Text == null) 
            {
                dcId = "";
            }  
            else 
            {
                dcId = this.dataCentre.Text;
            }
            
            HttpStatusCode redirectState;

            if (status302.Checked)
            {
                redirectState = HttpStatusCode.Redirect;
            } else if (statusOk.Checked) {
                   redirectState = HttpStatusCode.OK;
                } else {
                       redirectState = HttpStatusCode.MovedPermanently;
                    }

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
                string resultFile = string.Format(crNumber+"-"+dcId+"-Redirect-TestResults-{0:yyyy-MM-dd_hh-mm-ss-tt}.xls",
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
                xlWorkSheetNew.Cells[1, 4] = "EXPECTED HTTP RESPONSE CODE";
                xlWorkSheetNew.Cells[1, 5] = "ACTUAL HTTP RESPONSE CODE";
                xlWorkSheetNew.Cells[1, 6] = "RESULT";
           
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
                            Assert.AreEqual(redirectState, response.StatusCode);
                            StringAssert.AreEqualIgnoringCase(exp,act);
                            xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                            xlWorkSheetNew.Cells[rCnt, 1] = src;
                            xlWorkSheetNew.Cells[rCnt, 2] = exp;
                            xlWorkSheetNew.Cells[rCnt, 3] = act;
                            xlWorkSheetNew.Cells[rCnt, 4] = redirectState.ToString();
                            xlWorkSheetNew.Cells[rCnt, 5] = response.StatusCode.ToString();
                            xlWorkSheetNew.Cells[rCnt, 6] = "*****PASSED*****";
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
                            } 
 */

                        }
                        catch (AssertionException AE)
                        {
                            xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                            xlWorkSheetNew.Cells[rCnt, 1] = src;
                            xlWorkSheetNew.Cells[rCnt, 2] = exp;
                            xlWorkSheetNew.Cells[rCnt, 3] = act;
                            xlWorkSheetNew.Cells[rCnt, 4] = redirectState.ToString();
                            xlWorkSheetNew.Cells[rCnt, 5] = response.StatusCode.ToString();
                            xlWorkSheetNew.Cells[rCnt, 6] = "*****FAILED*****";
                            if (MessageBox.Show("********************ERROR********************"
                                + "\nROW: "
                                + rCnt
                                + "\nSOURCE URL: "
                                + src
                                + "\nEXPECTED DESTINATION URL: "
                                + exp
                                + "\nACTUAL DESTINATION URL: "
                                + act
                                + "\nEXPECTED HTTP RESPONSE CODE: "
                                + redirectState.ToString() 
                                + "\nACTUAL HTTP RESPONSE CODE: "
                                + response.StatusCode.ToString() 
                                /*+ "\nLOCATION: "  
                                + response.Headers.Location.ToString() 
                                + "\nREASON: "
                                + response.ReasonPhrase.ToString()
                                + "\nCONNECTION: "
                                + response.Headers.Connection*/
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

        private void radioButton1_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void label3_Click(object sender, EventArgs e)
        {

        }
    }
}
