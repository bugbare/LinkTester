using NUnit.Framework;

using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Chrome;


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
    public partial class BugBareSmoke : Form
    {

        private FirefoxDriver driver;
        //private ChromeDriver cDriver;
       

        /* Initialise the implemented Form, Form1 interface object */
        public BugBareSmoke()
        {
            InitializeComponent();
            driver = new FirefoxDriver();
            driver.Manage().Timeouts().ImplicitlyWait(new TimeSpan(0, 0, 10));
        }

        private void CreateButton_Click(object sender, EventArgs e)
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

            if (this.InputPageUrl.Text == "") 
                {
                    pageUrl = "http://www.cityindex.co.uk";
                }
            else 
                {
                    pageUrl = this.InputPageUrl.Text;
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
            xlWorkSheet.Name = string.Format("TestData-{0:yyyy-MM-dd}", DateTime.Today);
            
            /* Initialise a new Firefox driver object and go to the initial landing page*/
            

            /* Setup a headless http client handler that does not allow autoredirects, 
             * so we get the initial response status code returned by the server*/

            HttpClientHandler httpClientHandler = new HttpClientHandler();
            httpClientHandler.AllowAutoRedirect = true;


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

                /* Create a search pattern to filter attributes by content containing http*/
                string lPattern = "(http\\s?)";
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
                        var actualStatusCode = client.GetAsync(link).Result.StatusCode;
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
                        xlWorkSheet.Cells[rCnt, 3] = actualStatusCode;
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
                    driver.Manage().Cookies.DeleteAllCookies();
                    driver.Close();
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
            driver.Manage().Cookies.DeleteAllCookies();
            driver.Close();
            driver.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            MessageBox.Show("Excel file created , you can find the file at z:\\"+resultFile);
        }

        private void TestButton_Click(object sender, EventArgs e)
        {

            DialogResult result = readFileDialog.ShowDialog(); // Show the dialog.
            if (result == DialogResult.OK) // Test result.
            {
                string file = readFileDialog.FileName;

                Excel.Application xlApp;
                Excel.Workbook xlWorkBook;
                Excel.Worksheet xlWorkSheet;
                Excel.Range range;


                HttpClientHandler httpClientHandler = new HttpClientHandler();
                httpClientHandler.AllowAutoRedirect = true;

                HttpResponseMessage response;

                string sourceUrl;
                string expectedUrl;
                string actualUrl;
                string expectedTitle;
                string actualTitle;
                int expectedStatusCode;
                int actualStatusCode;
                int rCnt = 0;
                string resultFile = string.Format("Smoke-TestResults-{0:yyyy-MM-dd_hh-mm-ss-tt}.xls",
                                DateTime.Now);

                xlApp = new Excel.Application();
                xlWorkBook = xlApp.Workbooks.Open(file, 0, true, 5, "", "", true, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "\t", false, false, 0, true, 1, 0);
                xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);
                //driver = new FirefoxDriver();


                range = xlWorkSheet.UsedRange;
                int tRows = range.Rows.Count - 1;
                int tCols = range.Columns.Count;

                if (MessageBox.Show("TOTAL NUMBER OF ROWS TO TEST: " + tRows + "\n"
                    + "ROWS: "
                    + tRows
                    + "\nCOLUMNS: "
                    + tCols
                    , "TOTAL NUMBER OF TESTS",
                    MessageBoxButtons.OKCancel,
                    MessageBoxIcon.Asterisk,
                    MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                {
                    releaseObject(xlWorkSheet);
                    releaseObject(xlWorkBook);
                    releaseObject(xlApp);

                    driver.Manage().Cookies.DeleteAllCookies();
                    driver.Close();
                    driver.Quit();

                    Application.Exit();

                }

                Excel.Worksheet xlWorkSheetResult;
                xlWorkSheetResult = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                object misValue = System.Reflection.Missing.Value;
                xlWorkSheetResult.Name = string.Format("TestResults-{0:yyyy-MM-dd}", DateTime.Today);
                xlWorkSheetResult.Cells[1, 1] = "SOURCE URL ";
                xlWorkSheetResult.Cells[1, 2] = "EXPECTED DESTINATION URL";
                xlWorkSheetResult.Cells[1, 3] = "ACTUAL DESTINATION URL";
                xlWorkSheetResult.Cells[1, 4] = "EXPECTED HTTP RESPONSE CODE";
                xlWorkSheetResult.Cells[1, 5] = "ACTUAL HTTP RESPONSE CODE";
                xlWorkSheetResult.Cells[1, 6] = "EXPECTED PAGE TITLE";
                xlWorkSheetResult.Cells[1, 7] = "ACTUAL PAGE TITLE";
                xlWorkSheetResult.Cells[1, 8] = "RESULT";

                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {

                    sourceUrl = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                    string sPattern = "(http\\s?)";



                    if (System.Text.RegularExpressions.Regex.IsMatch(sourceUrl, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    {


                        HttpClient client = new HttpClient(httpClientHandler);
                        try
                        {
                            response = client.GetAsync(sourceUrl).Result;
                            actualStatusCode = (int)response.StatusCode;
                            driver.Navigate().GoToUrl(sourceUrl);
                            expectedUrl = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                            actualUrl = driver.Url.ToString();
                            expectedStatusCode = (int)(range.Cells[rCnt, 3] as Excel.Range).Value2;
                            expectedTitle = (string)(range.Cells[rCnt, 4] as Excel.Range).Value2;
                            actualTitle = driver.Title;

                            try
                            {
                                StringAssert.AreEqualIgnoringCase(expectedUrl, actualUrl);
                                Assert.AreEqual(expectedStatusCode, actualStatusCode);
                                StringAssert.AreEqualIgnoringCase(expectedTitle, actualTitle);
                                xlWorkSheetResult = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                                xlWorkSheetResult.Cells[rCnt, 1] = sourceUrl;
                                xlWorkSheetResult.Cells[rCnt, 2] = expectedUrl;
                                xlWorkSheetResult.Cells[rCnt, 3] = actualUrl;
                                xlWorkSheetResult.Cells[rCnt, 4] = expectedStatusCode.ToString();
                                xlWorkSheetResult.Cells[rCnt, 5] = actualStatusCode.ToString();
                                xlWorkSheetResult.Cells[rCnt, 6] = expectedTitle;
                                xlWorkSheetResult.Cells[rCnt, 7] = actualTitle;
                                xlWorkSheetResult.Cells[rCnt, 8] = "PASSED";
                            }
                            catch (AssertionException AE)
                            {
                                xlWorkSheetResult = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(2);
                                xlWorkSheetResult.Cells[rCnt, 1] = sourceUrl;
                                xlWorkSheetResult.Cells[rCnt, 2] = expectedUrl;
                                xlWorkSheetResult.Cells[rCnt, 3] = actualUrl;
                                xlWorkSheetResult.Cells[rCnt, 4] = expectedStatusCode.ToString();
                                xlWorkSheetResult.Cells[rCnt, 5] = actualStatusCode.ToString();
                                xlWorkSheetResult.Cells[rCnt, 6] = expectedTitle;
                                xlWorkSheetResult.Cells[rCnt, 7] = actualTitle;
                                
                                if (MessageBox.Show("********************IS THIS A FAILURE ???********************"
                                          + "\nROW: "
                                          + rCnt
                                          + "\nSOURCE URL: "
                                          + sourceUrl
                                          + "\nEXPECTED DESTINATION URL: "
                                          + expectedUrl
                                          + "\nACTUAL DESTINATION URL: "
                                          + actualUrl
                                          + "\nEXPECTED HTTP RESPONSE CODE: "
                                          + expectedStatusCode.ToString()
                                          + "\nACTUAL HTTP RESPONSE CODE: "
                                          + actualStatusCode.ToString()
                                          + "\nEXPECTED DESTINATION PAGE TITLE: "
                                          + expectedTitle
                                          + "\nACTUAL DESTINATION PAGE TITLE: "
                                          + actualTitle
                                          + "\n============================================="
                                          + "\nNUNIT Says: "
                                          + "\n"
                                          + AE.ToString(),
                                          "*****FAILED*****",
                                          MessageBoxButtons.YesNoCancel,
                                          MessageBoxIcon.Error,
                                          MessageBoxDefaultButton.Button1) == DialogResult.Yes)
                                {
                                    xlWorkSheetResult.Cells[rCnt, 8] = "*****FAILED*****";
                                }
                                else if (MessageBox.Show("********************IS THIS A FAILURE ???********************"
                                        + "\nROW: "
                                        + rCnt
                                        + "\nSOURCE URL: "
                                        + sourceUrl
                                        + "\nEXPECTED DESTINATION URL: "
                                        + expectedUrl
                                        + "\nACTUAL DESTINATION URL: "
                                        + actualUrl
                                        + "\nEXPECTED HTTP RESPONSE CODE: "
                                        + expectedStatusCode.ToString()
                                        + "\nACTUAL HTTP RESPONSE CODE: "
                                        + actualStatusCode.ToString()
                                        + "\nEXPECTED DESTINATION PAGE TITLE: "
                                        + expectedTitle
                                        + "\nACTUAL DESTINATION PAGE TITLE: "
                                        + actualTitle
                                        + "\n============================================="
                                        + "\nNUNIT Says: "
                                        + "\n"
                                        + AE.ToString(),
                                        "*****FAILED*****",
                                        MessageBoxButtons.YesNoCancel,
                                        MessageBoxIcon.Error,
                                        MessageBoxDefaultButton.Button1) == DialogResult.No)
                                {
                                    
                                    xlWorkSheetResult.Cells[rCnt, 8] = "*****PASSED (User Override)*****";
                                                                       
                                }
                                else if (MessageBox.Show("********************IS THIS A FAILURE ???********************"
                                    + "\nROW: "
                                    + rCnt
                                    + "\nSOURCE URL: "
                                    + sourceUrl
                                    + "\nEXPECTED DESTINATION URL: "
                                    + expectedUrl
                                    + "\nACTUAL DESTINATION URL: "
                                    + actualUrl
                                    + "\nEXPECTED HTTP RESPONSE CODE: "
                                    + expectedStatusCode.ToString()
                                    + "\nACTUAL HTTP RESPONSE CODE: "
                                    + actualStatusCode.ToString()
                                    + "\nEXPECTED DESTINATION PAGE TITLE: "
                                    + expectedTitle
                                    + "\nACTUAL DESTINATION PAGE TITLE: "
                                    + actualTitle
                                    + "\n============================================="
                                    + "\nNUNIT Says: "
                                    + "\n"
                                    + AE.ToString(),
                                    "*****FAILED*****",
                                    MessageBoxButtons.YesNoCancel,
                                    MessageBoxIcon.Error,
                                    MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                                {
                                    xlWorkSheetResult.Cells[rCnt, 8] = "*****FAILED*****";
                                    xlWorkBook.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                                    xlWorkBook.Close(true, misValue, misValue);
                                    xlApp.Quit();

                                    driver.Manage().Cookies.DeleteAllCookies();
                                    driver.Close();
                                    driver.Quit();

                                    releaseObject(xlWorkSheet);
                                    releaseObject(xlWorkSheetResult);
                                    releaseObject(xlWorkBook);
                                    releaseObject(xlApp);

                                    Application.Exit();
                                }

                            }

                        }

                        catch (Exception Timeout)
                        {
                            xlWorkSheetResult.Cells[rCnt, 1] = sourceUrl;
                            xlWorkSheetResult.Cells[rCnt, 2] = (range.Cells[rCnt, 2] as Excel.Range).Value2;
                            xlWorkSheetResult.Cells[rCnt, 3] = driver.Url;
                            xlWorkSheetResult.Cells[rCnt, 4] = (range.Cells[rCnt, 3] as Excel.Range).Value2;
                            xlWorkSheetResult.Cells[rCnt, 5] = client.GetAsync(sourceUrl).Result.StatusCode;
                            xlWorkSheetResult.Cells[rCnt, 6] = (range.Cells[rCnt, 4] as Excel.Range).Value2;
                            xlWorkSheetResult.Cells[rCnt, 7] = driver.Title;
                            xlWorkSheetResult.Cells[rCnt, 8] = "*****TIMEOUT*****";
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


                MessageBox.Show("TESTS COMPLETED");
                xlWorkBook.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBook.Close(true, misValue, misValue);

                MessageBox.Show("Excel file created , you can find the file at z:\\" + resultFile);


                xlApp.Quit();

                driver.Manage().Cookies.DeleteAllCookies();
                driver.Close();
                driver.Quit();

                releaseObject(xlWorkSheet);
                releaseObject(xlWorkSheetResult);
                releaseObject(xlWorkBook);
                releaseObject(xlApp);


                Application.Exit();
            }
        }

        /* Read Excel Spreadsheet  */
        private void ReadButton_Click(object sender, EventArgs e)
        {

            string crNumber;
            string dcId;

            if (this.ChangeRequest.Text == null) 
                {
                    crNumber = "";
                }
            else 
                {
                    crNumber = this.ChangeRequest.Text;
                }
             
            if (this.DataCentre.Text == null) 
            {
                dcId = "";
            }  
            else 
            {
                dcId = this.DataCentre.Text;
            }
            
            HttpStatusCode redirectState;

            if (Status302.Checked)
            {
                redirectState = HttpStatusCode.Redirect;
            } else if (StatusOk.Checked) {
                   redirectState = HttpStatusCode.OK;
                } else {
                       redirectState = HttpStatusCode.MovedPermanently;
                    }

            
            DialogResult result = readFileDialog.ShowDialog(); // Show the dialog.
	        if (result == DialogResult.OK) // Test result.
	        {
		        string file = readFileDialog.FileName;       

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
               // driver = new FirefoxDriver();

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

                    driver.Manage().Cookies.DeleteAllCookies();
                    driver.Close();
                    driver.Quit();
                }

                Excel.Workbook xlWorkBookNew;
                Excel.Worksheet xlWorkSheetNew;
                object misValue = System.Reflection.Missing.Value;
                xlWorkBookNew = xlApp.Workbooks.Add(misValue);
                xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                xlWorkSheetNew.Cells[1, 1] = "SOURCE URL";
                xlWorkSheetNew.Cells[1, 2] = "EXPECTED DESTINATION URL";
                xlWorkSheetNew.Cells[1, 3] = "ACTUAL DESTINATION URL";
                xlWorkSheetNew.Cells[1, 4] = "EXPECTED HTTP RESPONSE CODE";
                xlWorkSheetNew.Cells[1, 5] = "ACTUAL HTTP RESPONSE CODE";
                xlWorkSheetNew.Cells[1, 6] = "RESULT";
           
                for (rCnt = 1; rCnt <= range.Rows.Count; rCnt++)
                {

                    src = (string)(range.Cells[rCnt, 1] as Excel.Range).Value2;
                    string sPattern = "(http\\s?)";
                    

                        
                   if (System.Text.RegularExpressions.Regex.IsMatch(src, sPattern, System.Text.RegularExpressions.RegexOptions.IgnoreCase))
                    { 


                        HttpClient client = new HttpClient(httpClientHandler);
                        response = client.GetAsync(src).Result;

                        try
                        {
                            driver.Navigate().GoToUrl(src);
                            driver.Manage().Window.Maximize();
                            exp = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                            act = driver.Url.ToString();
                            try
                            {
                                Assert.AreEqual(redirectState, response.StatusCode);
                                StringAssert.AreEqualIgnoringCase(exp, act);
                                xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                                xlWorkSheetNew.Cells[rCnt, 1] = src;
                                xlWorkSheetNew.Cells[rCnt, 2] = exp;
                                xlWorkSheetNew.Cells[rCnt, 3] = act;
                                xlWorkSheetNew.Cells[rCnt, 4] = redirectState.ToString();
                                xlWorkSheetNew.Cells[rCnt, 5] = response.StatusCode.ToString();
                                xlWorkSheetNew.Cells[rCnt, 6] = "PASSED";
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

                                    driver.Manage().Cookies.DeleteAllCookies();
                                    driver.Close();
                                    driver.Quit();

                                    releaseObject(xlWorkSheet);
                                    releaseObject(xlWorkBook);
                                    releaseObject(xlWorkSheetNew);
                                    releaseObject(xlWorkBookNew);
                                    releaseObject(xlApp);

                                    Application.Exit();
                                }
                            }
                        }
                        catch (WebDriverException timeout)
                        {
                            exp = (string)(range.Cells[rCnt, 2] as Excel.Range).Value2;
                            act = driver.Url.ToString();

                            xlWorkSheetNew = (Excel.Worksheet)xlWorkBookNew.Worksheets.get_Item(1);
                            xlWorkSheetNew.Cells[rCnt, 1] = src;
                            xlWorkSheetNew.Cells[rCnt, 2] = exp;
                            xlWorkSheetNew.Cells[rCnt, 3] = act;
                            xlWorkSheetNew.Cells[rCnt, 4] = redirectState.ToString();
                            xlWorkSheetNew.Cells[rCnt, 5] = response.StatusCode.ToString();
                            xlWorkSheetNew.Cells[rCnt, 6] = "*****TIMEOUT*****";
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
                                + "\n============================================="
                                + "\nNUNIT Says: "
                                + "\n"
                                + timeout.ToString(),
                                "*****FAILED*****",
                                MessageBoxButtons.OKCancel,
                                MessageBoxIcon.Error,
                                MessageBoxDefaultButton.Button1) == DialogResult.Cancel)
                            {
                                xlWorkBook.Close(true, null, null);
                                xlWorkBookNew.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                                xlWorkBookNew.Close(true, misValue, misValue);
                                xlApp.Quit();

                                driver.Manage().Cookies.DeleteAllCookies();
                                driver.Close();
                                driver.Quit();

                                releaseObject(xlWorkSheet);
                                releaseObject(xlWorkBook);
                                releaseObject(xlWorkSheetNew);
                                releaseObject(xlWorkBookNew);
                                releaseObject(xlApp);

                                Application.Exit();
                            }
                        }
                    }                    
                }


                MessageBox.Show("TESTS COMPLETED");
                xlWorkBook.Close(true, null, null);
                xlWorkBookNew.SaveAs(resultFile, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
                xlWorkBookNew.Close(true, misValue, misValue);


                xlApp.Quit();

                driver.Manage().Cookies.DeleteAllCookies();
                driver.Close();
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

        private void BugBareSmoke_LoadFile(object sender, EventArgs e)
        {

        }

        private void RedirectMode_CheckBox_Click(object sender, EventArgs e)
        {

        }

        private void TestBuilder_Description_Click(object sender, EventArgs e)
        {
            MessageBox.Show("If you want to run a sample test - try entering 'http://www.cityindex.com'", "***** HINT *****", 
                MessageBoxButtons.OK, 
                MessageBoxIcon.Information,
                MessageBoxDefaultButton.Button1, 
                MessageBoxOptions.DefaultDesktopOnly);
        }

        private void Exit_BugBareSmoke(object sender, EventArgs e)
        {
            Application.Exit();
        }

    }
}
