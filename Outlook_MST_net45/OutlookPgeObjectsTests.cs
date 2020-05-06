
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using WinAppDriverPgeObjects.Views;
using System.Threading;
using OpenQA.Selenium.Support.UI;
using System.Diagnostics;

namespace WinAppDriverPgeObjects
{
    [TestClass]
    public class OutLookPageObjectsTests
    {
        // Note: append /wd/hub to the URL if you're directing the test at Appium
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string CalculatorAppId = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App";
        private const string OutLookAppId = "Outlook";
        //private const string OutLookAppId = @"Microsoft Outlook 2010";

        private static WindowsDriver<WindowsElement> _driver;
        private static OutLookStandardView _OutLookStandardView;
        private static WebDriverWait _waitEl;
        private static TestContext _testContext;
        public TestContext TestContext
        {
            get { return _testContext; }
            set { _testContext = value; }
        }

        public static string deployDir = "";
        [ClassInitialize]
        [DeploymentItem("MyFiles")]
        public static void ClassInit(TestContext context)
        {
            _testContext = context;
            deployDir = context.DeploymentDirectory;
            Console.WriteLine("ClassInitialize");
            AppiumOptions options = new AppiumOptions();
            options.AddAdditionalCapability("app", OutLookAppId);
            options.AddAdditionalCapability("deviceName", "WindowsPC");


            _driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), options);
            Assert.IsNotNull(_driver);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            _waitEl = new WebDriverWait(_driver, TimeSpan.FromSeconds(8));
            Assert.IsNotNull(_waitEl);

            // Identify the current window handle. You can check through inspect.exe which window this is.
            var currentWindowHandle = _driver.CurrentWindowHandle;

            // Wait for 5 seconds or however long it is needed for the right window to appear/for the splash screen to be dismissed
            Thread.Sleep(TimeSpan.FromSeconds(5));

            // Return all window handles associated with this process/application.
            // At this point hopefully you have one to pick from. Otherwise you can
            // simply iterate through them to identify the one you want.
            var allWindowHandles = _driver.WindowHandles;

            // Assuming you only have only one window entry in allWindowHandles and it is in fact the correct one,
            // switch the session to that window as follows. You can repeat this logic with any top window with the same
            // process id (any entry of allWindowHandles)
            _driver.SwitchTo().Window(allWindowHandles[0]);
            Console.WriteLine($"Application title: {_driver.Title}");
            _OutLookStandardView = new OutLookStandardView(_driver);
        }

        [ClassCleanup]
        public static void ClassCleanup()
        {
            Console.WriteLine("Console ClassCleanup");
            if (_driver != null)
            {
                _driver.Dispose();
                //_driver.Quit();
                _driver = null;
            }
        }

        [TestInitialize]
        public void BeforeClick()
        {
            Console.WriteLine("BeforeClick");
        }

        [TestCleanup]
        public void AfterClick()
        {
            Console.WriteLine("AfterClick");
        }

        [TestMethod]
        public void FolderClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            fName = TestContext.TestName;
            string txt = "";
            try
            {
                txt = _OutLookStandardView.FolderButton.Text;
                _OutLookStandardView.WaitUntil(_OutLookStandardView.FolderButton, _waitEl, fName);
                _OutLookStandardView.ClickElement(_OutLookStandardView.FolderButton, fName);
                txt = _OutLookStandardView.NewFolderButton.Text;
                _OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewFolderButton);
                //_OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewEmailButton);
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console " + $"catched: {fName}{exp.Message}");
                Assert.IsTrue(false, $"Testmetod {fName} failed.Element {txt} ");
            }
        }

        [TestMethod]
        public void HomeClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            fName = TestContext.TestName;
            string txt = "";
            try
            {
                txt = _OutLookStandardView.HomeButton.Text;
                _OutLookStandardView.ClickElement(_OutLookStandardView.HomeButton, fName);
                txt = _OutLookStandardView.NewEmailButton.Text;
                _OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewEmailButton);
                txt = "New E-mail1";
                WindowsElement el = _driver.FindElementByName("New E-mail1");
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console catched:{fName} {exp.Message}");
                Assert.IsTrue(false, $"Testmetod {fName} failed.Element {txt} ");
            }
            //var tvControl = _driver.FindElementByClassName("NetUIRibbonButton");

            //foreach (var tn in tvControl.FindElementsByClassName("NetUIRibbonButton"))
            //{
            //    Debug.WriteLine(
            //        $"*** BEFORE: {tn.Text} - Displayed: {tn.Displayed} - Enabled: {tn.Enabled} - Selected: {tn.Selected}");
            //    Console.WriteLine(
            //        $"*** BEFORE: {tn.Text} - Displayed: {tn.Displayed} - Enabled: {tn.Enabled} - Selected: {tn.Selected}");
            //}

            //_OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewFolderButton);
        }

        [TestMethod]
        public void ViewClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            fName = TestContext.TestName;
            string txt = "";
            try
            {
                txt = _OutLookStandardView.ViewButton.Text;
                _OutLookStandardView.ClickElement(_OutLookStandardView.ViewButton, fName);
                txt = _OutLookStandardView.ChangeViewButton.Text;
                _OutLookStandardView.AssertDisplayed(_OutLookStandardView.ChangeViewButton);
                _OutLookStandardView.ClickElement(_OutLookStandardView.ChangeViewButton, fName);
                txt = _OutLookStandardView.CompactButton.Text;
                _OutLookStandardView.CompactButton.Click();
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                txt = _OutLookStandardView.ChangeViewButton.Text;
                _OutLookStandardView.ClickElement(_OutLookStandardView.ChangeViewButton, fName);
                txt = _OutLookStandardView.SingleButton.Text;
                _OutLookStandardView.SingleButton.Click();
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                txt = _OutLookStandardView.ChangeViewButton.Text;
                _OutLookStandardView.ClickElement(_OutLookStandardView.ChangeViewButton, fName);
                txt = _OutLookStandardView.PreviewButton.Text;
                _OutLookStandardView.PreviewButton.Click();
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console catched:{fName} {exp.Message}");
                Assert.IsTrue(false, $"Testmetod {fName} failed.Element {txt} ");
            }
        }

        [TestMethod]
        public void ShowAsConversationsClick()
        {
            string txt = "";
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            fName = TestContext.TestName;
            try
            {
                txt = _OutLookStandardView.ViewButton.Text;
                _OutLookStandardView.ClickElement(_OutLookStandardView.ViewButton, fName);
                Console.WriteLine($"Value of checkbox is: {_OutLookStandardView.ShowAsConversationsButton.Selected}");
                if (_OutLookStandardView.ShowAsConversationsButton.Enabled)
                {
                    _OutLookStandardView.ClickElement(_OutLookStandardView.ShowAsConversationsButton, fName);
                    _OutLookStandardView.ThisFolderButton.Click();
                    //_OutLookStandardView.ClickElement(_OutLookStandardView.ThisFolderButton, fName);
                }
                Console.WriteLine($"Value of checkbox is: {_OutLookStandardView.ShowAsConversationsButton.Selected}");
                if (_OutLookStandardView.ShowAsConversationsButton.Enabled)
                {
                    _OutLookStandardView.ClickElement(_OutLookStandardView.ShowAsConversationsButton, fName);
                    _OutLookStandardView.ThisFolderButton.Click();
                    //_OutLookStandardView.ClickElement(_OutLookStandardView.ThisFolderButton, fName);
                }
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console catched:{fName} {exp.Message}");
                Assert.IsTrue(false, $"Testmetod {fName} failed.Element {txt} ");
            }
        }
    }
}
