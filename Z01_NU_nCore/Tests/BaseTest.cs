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
using Z01_NU_nCore.Pages;

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
        private static BasePage _BasePage;
        //private static WebDriverWait _waitEl;

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

            Thread.Sleep(TimeSpan.FromSeconds(8));
            wait = new WebDriverWait(driver, TimeSpan.FromSeconds(30));
            // _waitEl = new WebDriverWait(driver, TimeSpan.FromSeconds(8));
            Assert.IsNotNull(wait);
            helper.waitUntilOutlookOpens(driver);

            // Return all window handles associated with this process/application.
            var allWindowHandles = driver.WindowHandles;
            // Assuming you only have only one window entry in allWindowHandles and it is in fact the correct one,
            // switch the session to that window as follows. You can repeat this logic with any top window with the same
            // process id (any entry of allWindowHandles)
            driver.SwitchTo().Window(allWindowHandles[0]);

            _BasePage = new BasePage(driver);

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
 
        [Test]
        public void FolderClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            fName = TestContext.CurrentContext.Test.MethodName;
            string txt = "";
            try
            {
                txt = _BasePage.FolderButton.Text;
                _BasePage.WaitUntil(_BasePage.FolderButton, wait, fName);
                _BasePage.ClickElement(_BasePage.FolderButton, fName);
                txt = _BasePage.NewFolderButton.Text;
                _BasePage.AssertDisplayed(_BasePage.NewFolderButton);
                //_OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewEmailButton);
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console " + $"catched: {fName}{exp.Message}");
                Assert.IsTrue(false, $"Testmetod {fName} failed.Element {txt} ");
            }
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
                ////helper.killZeroProcess();
                string appdatapath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                ////helper.ZipDirectory(appdatapath + "\\ZeroOutlookAddin");
                ////helper.startZeroProcess();
                string path = Path.GetDirectoryName(BaseTest.reportPath) + "\\Logs";
                if (!Directory.Exists(path))
                {
                    Directory.CreateDirectory(path);
                }
                string zipfolderpath = path + "\\" + Guid.NewGuid().ToString() + ".zip";
                ////File.Move(appdatapath + "\\ZeroOutlookAddin.zip", zipfolderpath);
                extentTest.Fail("Run logs location: <a class='ziplink' href='" + zipfolderpath + "'>link</a>");
            }
            //  driver.CloseApp();
            //GlobalTestSetup.extent.Flush();
        }

        [OneTimeTearDown]
        public void OneTimeTearDown()
        {
            //GlobalTestSetup.extent.Flush();
            if (helper.isAppRuning(driver))
            {
                helper.closeAllApplevels(driver);
            }
            driver.Quit();
        }
    }
}
