
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

        [ClassInitialize]
        public static void ClassInit(TestContext context)
        {
            Debug.WriteLine("ClassInit");
            Console.WriteLine("ClassInitialize");
            AppiumOptions options = new AppiumOptions();
            options.AddAdditionalCapability("app", OutLookAppId);
            options.AddAdditionalCapability("deviceName", "WindowsPC");


            _driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), options);
            Assert.IsNotNull(_driver);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

            _waitEl = new WebDriverWait(_driver, TimeSpan.FromSeconds(5));
            Assert.IsNotNull(_waitEl);

            // Identify the current window handle. You can check through inspect.exe which window this is.
            //var currentWindowHandle = _driver.CurrentWindowHandle;

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
            Debug.WriteLine("Debug ClassCleanup");
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
            Debug.WriteLine("BeforeClick");
            Console.WriteLine("BeforeClick");
        }

        [TestCleanup]
        public void AfterClick()
        {
            Debug.WriteLine("AfterClick");
            Console.WriteLine("AfterClick");
        }

        [TestMethod]
        public void FolderClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                _OutLookStandardView.WaitUntil(_OutLookStandardView.FolderButton, _waitEl, fName);
                _OutLookStandardView.ClickElement(_OutLookStandardView.FolderButton, fName);
                _OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewFolderButton);
                //_OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewEmailButton);
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console catched: {fName}{exp.Message}");
                Debug.WriteLine($"Debug catched:{fName} {exp.Message}");
            }
        }

        [TestMethod]
        public void HomeClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                _OutLookStandardView.ClickElement(_OutLookStandardView.HomeButton, fName);
                string txt = _OutLookStandardView.NewEmailButton.Text;
                _OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewEmailButton);
                Debug.WriteLine($"Debug: Test: {fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                WindowsElement el = _driver.FindElementByName("New Email");
                Debug.WriteLine($"Debug: Test: {fName} {"after New Email"} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                Console.WriteLine($"Console: Test:{fName} {"after New Email"} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console catched:{fName} {exp.Message} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                Debug.WriteLine($"Debug catched:{fName} {exp.Message} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                //throw new Exception($"Testmetod {fName} failed ");
                Assert.IsTrue(false);
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
            try
            {
                string txt = "";
                _OutLookStandardView.ClickElement(_OutLookStandardView.ViewButton, fName);
                _OutLookStandardView.AssertDisplayed(_OutLookStandardView.ChangeViewButton);
                _OutLookStandardView.ClickElement(_OutLookStandardView.ChangeViewButton, fName);
                txt = _OutLookStandardView.CompactButton.Text;
                _OutLookStandardView.CompactButton.Click();
                Debug.WriteLine($"Debug: Test: {fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                //_OutLookStandardView.WaitUntil(_OutLookStandardView.CompactButton, _waitEl, fName);
                //_OutLookStandardView.ClickElement(_OutLookStandardView.CompactButton, fName);
                _OutLookStandardView.ClickElement(_OutLookStandardView.ChangeViewButton, fName);
                txt = _OutLookStandardView.SingleButton.Text;
                _OutLookStandardView.SingleButton.Click();
                Debug.WriteLine($"Debug: Test: {fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                //_OutLookStandardView.ClickElement(_OutLookStandardView.SingleButton, fName);
                _OutLookStandardView.ClickElement(_OutLookStandardView.ChangeViewButton, fName);
                txt = _OutLookStandardView.PreviewButton.Text;
                _OutLookStandardView.PreviewButton.Click();
                Debug.WriteLine($"Debug: Test: {fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                Console.WriteLine($"Console: Test:{fName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
                //_OutLookStandardView.WaitUntil(_OutLookStandardView.PreviewButton, _waitEl, fName);
                //_OutLookStandardView.ClickElement(_OutLookStandardView.PreviewButton, fName);
            }
            catch (Exception exp)
            {
                Console.WriteLine($"Console catched:{fName} {exp.Message}");
                Debug.WriteLine($"Debug catched:{fName} {exp.Message}");
                Assert.IsTrue(false);
            }
        }

        [TestMethod]
        public void ShowAsConversationsClick()
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            try
            {
                _OutLookStandardView.ClickElement(_OutLookStandardView.ViewButton, fName);
                Debug.WriteLine($"Value of checkbox is: {_OutLookStandardView.ShowAsConversationsButton.Selected}");
                Console.WriteLine($"Value of checkbox is: {_OutLookStandardView.ShowAsConversationsButton.Selected}");
                if (_OutLookStandardView.ShowAsConversationsButton.Enabled)
                {
                    _OutLookStandardView.ClickElement(_OutLookStandardView.ShowAsConversationsButton, fName);
                    _OutLookStandardView.ThisFolderButton.Click();
                    //_OutLookStandardView.ClickElement(_OutLookStandardView.ThisFolderButton, fName);
                }
                Debug.WriteLine($"Value of checkbox is: {_OutLookStandardView.ShowAsConversationsButton.Selected}");
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
                Debug.WriteLine($"Debug catched: {fName}{exp.Message}");
            }
        }
    }
}
