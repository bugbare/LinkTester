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
        /* Initialise the implemented Form, Form1 interface object */
        public Form1()
        {
            InitializeComponent();
        }

        private void createButton_Click(object sender, EventArgs e)
        {
            /* Instantiate Excel App object to use to generate an excel workbook*/
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();
            /* Creata a local String object to hold the pageUrl to be entered subsequently via the form objects*/
            String pageUrl;

            /* Check that the host system has excel installed, if not output an error message, with only an exit confirmation button*/

            if (xlApp == null)
            {
                if (MessageBox.Show("Excel is not properly installed!!", "No Excel Found - Application will Exit", MessageBoxButtons.OK,
                    MessageBoxIcon.Asterisk,
                    MessageBoxDefaultButton.Button1) == DialogResult.OK)
                {
                    Application.Exit();
                }
                return;
            }

            /* If there is no user input for the URL test field then default to City Index UK site */

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
             * to ignore certain options that are not required to be used - null is not allowed to be used! */

            object misValue = System.Reflection.Missing.Value;

            /* Instantiate a workbook and initialise the first worksheet, assigned to the xlWorksheet object*/
            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
            
            /* Initialise a new webdriver driver object and go to the initial landing page*/
            IWebDriver driver = new FirefoxDriver();

            /* Setup a headless http client handler that does not allow autoredirects, 
             * so we get the initial response status code returned by the server*/

            HttpClientHandler httpClientHandler = new HttpClientHandler();
            httpClientHandler.AllowAutoRedirect = false;

            HttpResponseMessage response;

            /* Instantiate a new Http Client object and apply the attributes of the previously generated http client handler to it */
            HttpClient client = new HttpClient(httpClientHandler);
            /* Create a file */
            string resultFile = string.Format("Test Data-{0:yyyy-MM-dd_hh-mm-ss-tt}.xls",
                                DateTime.Now);

            try
            {            /* use firefox (driver) to navigate to the initial page sent by user or default value provided by app */
                driver.Navigate().GoToUrl(pageUrl);

                /* Create a local variable to hold contents of the page source*/
                var pageLinks = new List<string>();

                /* Get all attribute elements on the page and store in a locally declared variable */
                var anchors = driver.FindElements(By.TagName("a"));

                /* Add a header row to the first row in the excel worksheet*/
                xlWorkSheet.Cells[1, 1] = "SOURCE URL ";
                xlWorkSheet.Cells[1, 2] = "DESTINATION URL";
                xlWorkSheet.Cells[1, 3] = "HTTP RESPONSE CODE";
                xlWorkSheet.Cells[1, 4] = "DESTINATION PAGE TITLE";

                /* Rest the row count to 2, so that it will start adding entries after the header row in the excel worksheet */
                int rCnt = 2;

                /* Create a search pattern to filter attributes by content contianing http*/
                string lPattern = "http";
                foreach (var element in anchors)
                /* filter attributes further by their category = href */
                {
                    if (element.GetAttribute("href") != null)

                    /* store the value of each href attribute into a pagelink list string object */
                    {
                        string pageLink = element.GetAttribute("href");

                        /* Filter the list objects by the http search pattern, so only legitimate http links are stored */
                        if (System.Text.RegularExpressions.Regex.IsMatch(pageLink, lPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                        {
                            /* Filter the list objects further to remove any that contain a script call "/#" */
                            if (System.Text.RegularExpressions.Regex.IsMatch(pageLink, "/#", System.Text.RegularExpressions.RegexOptions.IgnoreCase) == false)
                            {
                                pageLinks.Add(pageLink);
                                xlWorkSheet.Cells[rCnt, 1] = pageLink;
                                rCnt++;
                            }
                        }
                    }
                }

                /* Rest the row count to the second row on the excel sheet, so that values 
                 * that result from navigating to the value in the first cell of that row 
                 * Can be stored in proceeding columns in the same row */

                rCnt = 2;

                /* Use a foreach loop to cycle through all objects in the pageLinks list */
                foreach (var link in pageLinks)
                {

                    try
                    {
                        /* store the http response header retrieved from navigating using http 
                         * client to the http link provided by the current pageList object */
                        response = client.GetAsync(pageUrl).Result;
                        /* 
                         * DEBUG MessageBox.Show("Response Status headers: " + response.Headers, "STATUS CODE", MessageBoxButtons.OK); 
                         */

                        /* Navigate to the URL stored in the current pageList object */
                        driver.Navigate().GoToUrl(link);

                        /* retreive and store the current URL value */
                        string destinationUrl = driver.Url;

                        /* retrieve and store the current URL Page Title */
                        string pageTitle = driver.Title;

                        /* write out values retrieved after page navigation to excel worksheet in same row as source URL */
                        xlWorkSheet.Cells[rCnt, 2] = destinationUrl;
                        xlWorkSheet.Cells[rCnt, 3] = response.StatusCode;
                        xlWorkSheet.Cells[rCnt, 4] = pageTitle;

                        /* Return to source page */
                        //driver.Navigate().Back();
                        rCnt++;

                    }
                    catch (Exception Timeout)
                    {
                        xlWorkSheet.Cells[rCnt, 2] = driver.Url;
                        xlWorkSheet.Cells[rCnt, 3] = "TIMEOUT FAILURE";
                        xlWorkSheet.Cells[rCnt, 4] = driver.Title;
                        driver.Navigate().Back();
                        rCnt++;
                        MessageBox.Show(
                        Timeout.ToString()
                        + "\nFound before ROW: "
                        + rCnt.ToString(),
                        "********************TIMEOUT********************",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Error);
                    }
                }
            }
            catch (Exception ae)
            {
                if (MessageBox.Show(               
                    ae.ToString(),
                    "********************EXCEPTION********************",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Error,
                    MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                {
                    xlWorkBook.Close(true, null, null);
                    xlWorkBook.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                    xlWorkBook.Close(true, misValue, misValue);
                    xlApp.Quit();
                    driver.Quit();

                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    Application.Exit();
                }
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
                xlWorkSheetNew.Cells[1, 1] = "SOURCE URL ";
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
