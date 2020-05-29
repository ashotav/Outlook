using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using AventStack.ExtentReports.Reporter.Configuration;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Text;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Threading;
using System.Threading.Tasks;
using Z01_NU_nCore.Helpers;

namespace Z01_NU_nCore.Tests
{
    [TestFixture]
    public class BaseTest
    {
        private const string appPath = "Outlook";
        // Note: append /wd/hub to the URL if you're directing the test via Appium
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        public static WindowsDriver<WindowsElement> driver;
        //public static ExtentReports extent;
        public static ExtentTest extentTest;
        public static ExtentHtmlReporter htmlReporter;
        public Helper helper;
        public static WebDriverWait wait;
        public static string assemblyPath;
        public static string reportPath;

        [OneTimeSetUp]
        public void OneTimeSetUp()
        {
            reportPath = GlobalTestSetup.reportPath;
            helper = new Helper();
            assemblyPath = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
            // string projectPath = @"C:\Program Files (x86)\Jenkins\workspace\TestRunner";
            //helper.startZeroProcess();
            AppiumOptions opt = new AppiumOptions();
            opt.AddAdditionalCapability("app", appPath);
            opt.AddAdditionalCapability("deviceName", "WindowsPC");
            driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), opt, TimeSpan.FromMinutes(1));
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(40);
            Thread.Sleep(8000);
            wait = new WebDriverWait(driver, new TimeSpan(0, 0, 120));
            helper.waitUntilOutlookOpens(driver);
            var allWindowHandles = driver.WindowHandles;
            driver.SwitchTo().Window(allWindowHandles[0]);
        }

        [SetUp]
        public void Setup()
        {
            string testname = TestContext.CurrentContext.Test.MethodName;
            //helper.cleanupInbox();
            extentTest = GlobalTestSetup.extent.CreateTest(testname);
            if (!helper.isAppRuning(driver))
            {
                driver.LaunchApp();
                helper.waitUntilOutlookOpens(driver);
            }
            extentTest.Info("Outlook opens");
        }

        [Test]
        public void PlaceHolder()
        {

        }

        [TearDown]
        public void TearDown()
        {
            TestStatus status = TestContext.CurrentContext.Result.Outcome.Status;
            //ToDo Write exception in test results
            if (status == TestStatus.Failed)
            {
                var errorMessage = TestContext.CurrentContext.Result.Message;
                string stackTrace = TestContext.CurrentContext.Result.StackTrace.ToString();
                extentTest.Fail(errorMessage + "<br>" + stackTrace);
                helper.closeAllApplevels(driver);
                helper.killZeroProcess();
                string appdatapath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                helper.ZipDirectory(appdatapath + "\\ZeroOutlookAddin");
                helper.startZeroProcess();
                string path = Path.GetDirectoryName(BaseTest.reportPath) + "\\Logs";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string zipfolderpath = path + "\\" + Guid.NewGuid().ToString() + ".zip";
                File.Move(appdatapath + "\\ZeroOutlookAddin.zip", zipfolderpath);
                extentTest.Fail("Run logs location: <a class='ziplink' href='" + zipfolderpath + "'>link</a>");
            }
            //  driver.CloseApp();
            //  extent.Flush();
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            //   extent.Flush();
            if (helper.isAppRuning(driver))
            {
                helper.closeAllApplevels(driver);
            }
            driver.Quit();
        }
    }
}
