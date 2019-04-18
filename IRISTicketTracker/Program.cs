using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using Newtonsoft.Json;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Support.UI;
using Microsoft.Extensions.Configuration;
using NLog;
using System.Diagnostics;
using System.Threading.Tasks;
using JnJ.EAS.TicketTracker.Core;
using JnJ.EAS.TicketTracker.Core.Models;

namespace IRISTicketTracker
{
    class Program
    {
        private static IConfigurationRoot configuration;
        static List<TicketVolumn> lstAppBaselineYrly = new List<TicketVolumn>();
        static List<TicketVolumn> lstAppBaselineMonthly = new List<TicketVolumn>();
        static List<Report> lstSRReport = new List<Report>();
        static List<IncidentReport> lstIncReport = new List<IncidentReport>();
        static List<ChangeRequest> lstChanges = new List<ChangeRequest>();

        private static Logger logger = LogManager.GetLogger("fileLogger");
        private static Logger errorLogger = LogManager.GetLogger("errorFileLogger");

        static List<string> invalidIncCodes = new List<string>();

        static void Main(string[] args)
        {
            var stopWatch = Stopwatch.StartNew();
            logger.Info("Program started");

            ConfigurationHelper.LoadConfig();

            //let's get invalidIncCodes from app settings and populate the list object
            invalidIncCodes = ConfigurationHelper.GetAppSettingValue("InvalidIncidentResolutionCodes").Split(',').ToList();

            Parallel.Invoke(
                    () => LoadBaselineData(),
                    () => OpenIncidentReport(),
                    () => OpenServiceRequestReport(),
                    () => GetChangesRequests()
                );

            RunCalculations();

            Console.WriteLine(stopWatch.Elapsed.TotalMilliseconds);
            logger.Info("Elapsed time: " + stopWatch.Elapsed.TotalMilliseconds + " ms");
            logger.Info("==================================== Program Ended =========================================");
        }

        protected static void GetChangesRequests()
        {
            try
            {
                Console.WriteLine("Loading Change Request data.... please, wait...");
                logger.Info("Loading Change Request data.... please, wait...");

                SPUtility sPUtility = new SPUtility();
                lstChanges = sPUtility.ReadData();

                Console.WriteLine("Change Requests have been successfully retreived from SharePoint!");
                logger.Info("Change Requests have been successfully retreived from SharePoint!");
            }
            catch (Exception ex)
            {
                logger.Info("Failed at loading CR data. Error: " + ex.Message);
                errorLogger.Error(ex.Message);                
            }
        }
        public static void ReadDataFromExcel()
        {
            ExcelUtility xlUtil = new ExcelUtility();
            xlUtil.ReadData();
        }

        public static void LoadBaselineData()
        {
            try
            {
                Console.WriteLine("Loading baseline data.... please, wait...");
                logger.Info("Loading baseline data.... please, wait...");

                string json = File.ReadAllText(ConfigurationHelper.GetAppSettingValue("BaselineJSONPath"));
                lstAppBaselineYrly = JsonConvert.DeserializeObject<List<TicketVolumn>>(json);

                foreach (var item in lstAppBaselineYrly)
                {
                    var monthlyVol = new TicketVolumn();
                    monthlyVol.ApplicationID = item.ApplicationID;
                    monthlyVol.ApplicationName = item.ApplicationName;
                    monthlyVol.Group = item.Group;
                    monthlyVol.SLA = item.SLA;
                    monthlyVol.Incidents = item.Incidents / 12;
                    monthlyVol.ServiceRequest = item.ServiceRequest / 12;
                    monthlyVol.ChangeRequest = item.ChangeRequest / 12;
                    monthlyVol.Problem = item.Problem / 12;
                    monthlyVol.NonTicketed = item.NonTicketed / 12;
                    monthlyVol.Total = item.Total / 12;
                    monthlyVol.IsActive = item.IsActive;

                    lstAppBaselineMonthly.Add(monthlyVol);
                }

                Console.WriteLine("Baseline data has been loaded!");
                logger.Info("Baseline data has been loaded!");
            }
            catch (Exception ex)
            {
                logger.Info("Failed at loading baseline data. Error: " + ex.Message);
                errorLogger.Error(ex.Message);
            }
        }

        public static void OpenServiceRequestReport()
        {
            try
            {
                Console.WriteLine("Initiating the loading of Service Request data.... please, wait...");
                logger.Info("Initiating the loading of Service Request data.... please, wait...");

                var options = new ChromeOptions();
                options.AddArgument("headless");

                IWebDriver driver;
                driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options)
                {
                    Url = ConfigurationHelper.GetAppSettingValue("ServiceRequestReportUrl")
                };

                Console.WriteLine("Browsing IRIS SR report to pull Service Request data.... please, wait...");
                logger.Info("Browsing IRIS SR report to pull Service Request data.... please, wait...");
                IWebElement frame = (new WebDriverWait(driver, TimeSpan.FromMinutes(1)))
                    .Until(d => d.FindElement(By.Id("gsft_main")));

                driver.SwitchTo().Frame(frame);

                // let's build a list of tickets
                var tbody = (new WebDriverWait(driver, TimeSpan.FromMinutes(1)))
                        .Until(d => d.FindElement(By.ClassName("list2_body")));


                int rowCount = tbody.FindElements(By.TagName("tr")).Count;
                int colCount = (tbody.FindElements(By.TagName("tr"))[0]).FindElements(By.ClassName("vt")).Count;

                foreach (var rowItem in tbody.FindElements(By.TagName("tr")))
                {
                    var cells = rowItem.FindElements(By.ClassName("vt"));
                    Report report = new Report()
                    {
                        Numnber = cells[0].Text,
                        ApplicationID = cells[1].Text,
                        ApplicationName = cells[2].Text,
                        Priority = cells[3].Text,
                        State = cells[4].Text,
                        AssignedTo = cells[5].Text,
                        ShortDescription = cells[6].Text,
                        AssignmentGroup = cells[7].Text,
                        TaskType = cells[8].Text,
                        Opened = cells[9].Text,
                        CustTime = cells[10].Text,
                        Duration = cells[11].Text,
                        Status = cells[12].Text,
                        TaskState = cells[13].Text
                    };

                    lstSRReport.Add(report);
                }

                driver.Close();
                driver.Quit();

                Console.WriteLine("SR data has been loaded successfully!");
                logger.Info("SR data has been loaded successfully!");
            }
            catch (Exception ex)
            {
                logger.Info("Failed at loading SR data. Error: " + ex.Message);
                errorLogger.Error(ex.Message);
            }
        }

        public static void OpenIncidentReport()
        {
            try
            {
                Console.WriteLine("Initiating the loading of Incident tickets.... please, wait...");
                logger.Info("Initiating the loading of Incident tickets.... please, wait...");

                var options = new ChromeOptions();
                options.AddArgument("headless");

                IWebDriver driver;
                driver = new ChromeDriver(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), options)
                {
                    Url = ConfigurationHelper.GetAppSettingValue("IncidentReportUrl")
                };

                Console.WriteLine("Browsing IRIS report to pull Incident data.... please, wait...");
                logger.Info("Browsing IRIS report to pull Incident data.... please, wait...");

                IWebElement frame = (new WebDriverWait(driver, TimeSpan.FromMinutes(1)))
                    .Until(d => d.FindElement(By.Id("gsft_main")));

                driver.SwitchTo().Frame(frame);

                // let's build a list of tickets
                var tbody = (new WebDriverWait(driver, TimeSpan.FromMinutes(1)))
                        .Until(d => d.FindElement(By.ClassName("list2_body")));


                int rowCount = tbody.FindElements(By.TagName("tr")).Count;
                int colCount = (tbody.FindElements(By.TagName("tr"))[0]).FindElements(By.ClassName("vt")).Count;

                foreach (var rowItem in tbody.FindElements(By.TagName("tr")))
                {
                    var cells = rowItem.FindElements(By.ClassName("vt"));
                    IncidentReport report = new IncidentReport()
                    {
                        Numnber = cells[0].Text,
                        ApplicationID = cells[1].Text,
                        ApplicationName = cells[2].Text,
                        Priority = cells[3].Text,
                        ShortDescription = cells[4].Text,
                        State = cells[5].Text,
                        AssignedTo = cells[6].Text,
                        Opened = cells[7].Text,
                        Resolved = cells[8].Text,
                        Closed = cells[9].Text,
                        Category = cells[10].Text,
                        Status = cells[11].Text,
                        AssignmentGroup = cells[12].Text,
                        ResolutionCategory = cells[13].Text,
                        ResolutionCode = cells[14].Text,
                        KCS = cells[15].Text
                    };

                    lstIncReport.Add(report);
                }

                driver.Close();
                driver.Quit();
                Console.WriteLine("Incident data has been successfully loaded!");
                logger.Info("Incident data has been successfully loaded!");
            }
            catch (Exception ex)
            {
                logger.Info("Failed at loading INCIDENT data. Error: " + ex.Message);
                errorLogger.Error(ex.Message);
            }
        }

        public static void RunCalculations()
        {
            try
            {
                Console.WriteLine("Running calculations on the retrieved data to prepare final report.. please, wait...");
                logger.Info("Running calculations on the retrieved data to prepare final report.. please, wait...");

                StringBuilder html = new StringBuilder();
                html.Append("<table border=1 cellpadding=0 cellspacing=0 style='font-family: Verdana; font-size:9pt; padding: 5px'>");
                html.Append("<tr><th rowspan=2>AppID</th><th rowspan=2>Name</th><th rowspan=2>SLA</th><th colspan=2>Inc</th><th colspan=2>SR</th><th colspan=2>NT</th><th colspan=2>CR</th><th colspan=2>Prob</th><th colspan=2>Total</th></tr>");
                html.Append("<tr><th>Baseline</th><th>Actual</th><th>Baseline</th><th>Actual</th><th>Baseline</th><th>Actual</th><th>Baseline</th><th>Actual</th><th>Baseline</th><th>Actual</th><th>Baseline</th><th>Actual</th></tr><tbody>");

                //List<TicketVolumn> lstMonthlyTicketVol = new List<TicketVolumn>();
                //now we need to loop thru the disntict list and prepare ticket volumn stat
                foreach (var app in lstAppBaselineMonthly.Where(x => x.IsActive == true))
                {

                    var result = lstSRReport.Where(l => l.ApplicationID == app.ApplicationID);

                    var incCount = lstIncReport.Where(x => x.ApplicationID == app.ApplicationID
                                                        && !invalidIncCodes.Any(y => y.ToLower().Trim() == x.ResolutionCode.ToLower().Trim())).Count();

                    var srCount = result.Where(l => l.ShortDescription.ToLower().Trim() != "Operational Tasks".ToLower()
                                                    && l.Status.ToLower().Trim() != "Not Required".ToLower()).Count();

                    var ntCount = result.Where(l => l.ShortDescription.ToLower().Trim() == "Operational Tasks".ToLower()
                                                    && l.Status.ToLower().Trim() != "Not Required".ToLower()).Count();

                    // oh wait, we dont have AppID as a seaparate field in Change tracker, CR AppName has appid in it
                    var crCount = lstChanges.Where(x => x.Application.StartsWith(app.ApplicationID)
                                                        && string.IsNullOrEmpty(x.SOW)
                                                        && x.Created.Year == System.DateTime.Now.Year
                                                        && x.Created.Month == System.DateTime.Now.Month).Count();

                    html.Append("<tr>");

                    html.Append("<td>" + app.ApplicationID + "</td>");
                    html.Append("<td>" + app.ApplicationName + "</td>"); //app name
                    html.Append("<td>" + app.SLA + "</td>"); // sla
                    html.Append("<td>" + app.Incidents + "</td>"); // inc baseline
                    html.Append("<td>" + incCount + "</td>"); // inc actual
                    html.Append("<td>" + app.ServiceRequest + "</td>"); // sr baseline
                    html.Append("<td>" + srCount + "</td>");
                    html.Append("<td>" + app.NonTicketed + "</td>"); // NT baseline
                    html.Append("<td>" + ntCount + "</td>");
                    html.Append("<td>" + app.ChangeRequest + "</td>");
                    html.Append("<td>" + crCount + "</td>");
                    html.Append("<td>" + app.Problem + "</td>");
                    html.Append("<td>TBD</td>");
                    html.Append("<td>" + app.Total + "</td>");
                    html.Append("<td>" + (incCount + srCount + ntCount + crCount) + "</td>");
                    html.Append("</tr>");

                }
                html.Append("</tbody></table>");

                Console.WriteLine("Final report has been generated!");
                logger.Info("Final report has been generated!");

                if (ConfigurationHelper.GetAppSettingValue("EnableEmailNotification") == "true")
                {
                    Console.WriteLine("Sending report as email to the recepients.... please, wait... ");
                    logger.Info("Sending report as email to the recepients.... please, wait... ");

                    string mailBody = "<html><head></head><body>" + html.ToString() + "</body></html>";

                    MailUtility mailUtil = new MailUtility();
                    mailUtil.SendMail(ConfigurationHelper.GetAppSettingValue("MailFrom"), ConfigurationHelper.GetAppSettingValue("MailTo"), ConfigurationHelper.GetAppSettingValue("MailSubject"), mailBody);

                    Console.WriteLine("Mail sent!");
                    logger.Info("Mail sent!");
                }
                Console.WriteLine("Process completed!");
                logger.Info("Process completed!");
            }
            catch (Exception ex)
            {
                logger.Info("Something went wrong while running calculation. Error: " + ex.Message);
                errorLogger.Error(ex.Message);
            }
        }
    }

    
}
