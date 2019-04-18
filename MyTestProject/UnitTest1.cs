using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.IE;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Reflection;
using System.Threading;

namespace MyTestProject
{
    [TestClass]
    public class MyTestCases
    {
        [TestMethod]
        public void LoadUCT()
        {
            IWebDriver driver;
            driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))
            {
                Url = "http://uctqa.jnj.com"
            };
            driver.Close();
        }

        [TestMethod]
        public void OpenServiceRequestReport()
        {
            IWebDriver driver;
            driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location))
            {
                Url = "https://jnjprod.service-now.com/nav_to.do?uri=%2Fsys_report_template.do%3Fjvar_report_id%3D40322c24dbf87388e7e8a2aa489619a9%26jvar_selected_chart_options_tab%3D%26jvar_report_home_query%3D%26jvar_selected_tab%3DmyReports"
            };

            IWebElement frame = (new WebDriverWait(driver, TimeSpan.FromMinutes(1)))
                .Until(d => d.FindElement(By.Id("gsft_main")));

            driver.SwitchTo().Frame(frame);

            // let's build a list of tickets
            var tbody = (new WebDriverWait(driver, TimeSpan.FromMinutes(1)))
                    .Until(d => d.FindElement(By.ClassName("list2_body")));

            
            int rowCount = tbody.FindElements(By.TagName("tr")).Count;
            int colCount = (tbody.FindElements(By.TagName("tr"))[0]).FindElements(By.ClassName("vt")).Count;

            //ArrayList lstRecords = new ArrayList();
            DataTable dt = new DataTable();
            foreach (var rowItem in tbody.FindElements(By.TagName("tr")))
            {
                DataRow dr = dt.NewRow();
                foreach (var tdItem in rowItem.FindElements(By.ClassName("vt")))
                {
                    DataColumn dc = new DataColumn();                    
                    dt.Columns.Add(dc);
                    dr[dc] = tdItem.Text;
                }
            }
            //Console.WriteLine(rowCount);
            //Assert.AreEqual(0, colCount);
        }

        [TestMethod]
        public void RunReportOnIE()
        {
            var options = new InternetExplorerOptions()
            {
                InitialBrowserUrl = "https://jnjprod.service-now.com/nav_to.do?uri=%2Fsys_report_template.do%3Fjvar_report_id%3D5fcf83b5dba0338c6d0fdb41ca961910%26jvar_selected_tab%3DmyReports%26jvar_list_order_by%3D%26jvar_list_sort_direction%3D%26sysparm_reportquery%3D%26jvar_search_created_by%3D%26jvar_search_table%3D%26jvar_search_report_sys_id%3D%26jvar_report_home_query%3D",
                IntroduceInstabilityByIgnoringProtectedModeSettings = true
            };

            IWebDriver ieDriver 
                = new InternetExplorerDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options);

            Thread.Sleep(60000);
            IWebElement elem = ieDriver.FindElement(By.ClassName("collapsedGroup"));
            //ieDriver.Close();
            //ieDriver.Quit();
        }
    }
}
