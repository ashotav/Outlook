
using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Remote;
using oWinAppDriverPageObjects.Views;

namespace oWinAppDriverPageObjects
{
    [TestClass]
    public class OutLookPageObjectsTests
    {
        // Note: append /wd/hub to the URL if you're directing the test at Appium
        private const string WindowsApplicationDriverUrl = "http://127.0.0.1:4723";
        //private const string CalculatorAppId = "Microsoft.WindowsCalculator_8wekyb3d8bbwe!App";
        //private const string OutLookAppId = @"C:/Program Files/Microsoft Office/Office14/OUTLOOK.EXE";
        private const string OutLookAppId = "Outlook";

        private static WindowsDriver<WindowsElement> _driver;
        private static OutLookStandardView _OutLookStandardView;

        [ClassInitialize]
        public static void ClassInit(TestContext _ctx)
        {
            AppiumOptions options = new AppiumOptions();
            options.AddAdditionalCapability("app",OutLookAppId);
            options.AddAdditionalCapability("deviceName", "WindowsPC");

            _driver = new WindowsDriver<WindowsElement>(new Uri(WindowsApplicationDriverUrl), options);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(5);

           _OutLookStandardView = new OutLookStandardView(_driver);

        }

    [ClassCleanup]
        public static void ClassCleanup()
        {
            if (_driver != null)
            {
                _driver.Quit();
                _driver = null;
            }
        }

        [TestMethod]
        public void Clicking()
        {
            _driver.FindElementByName("Home").Click();
            _OutLookStandardView.ClickByTab(_OutLookStandardView.FolderButton);
            _driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(8);
            _OutLookStandardView.ClickByTab(_OutLookStandardView.HomeButton);
            //_OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewFolderButton);
            _OutLookStandardView.AssertDisplayed(_OutLookStandardView.NewE_MailButton);
            //_calcStandardView.AssertResult(12);
        }
        /*
                [TestMethod]
                public void Division()
                {
                    _calcStandardView.PerformCalculation(8, '/', 1);

                    _calcStandardView.AssertResult(8);
                }

                [TestMethod]
                public void Multiplication()
                {
                    _calcStandardView.PerformCalculation(9, '*', 9);

                    _calcStandardView.AssertResult(81);
                }

                [TestMethod]
                public void Subtraction()
                {
                    _calcStandardView.PerformCalculation(9, '-', 1);

                    _calcStandardView.AssertResult(8);
                }

                [TestMethod]
                [DataRow(1, '+', 7, 8)]
                [DataRow(9, '-', 7, 2)]
                [DataRow(8, '/', 4, 2)]
                public void FewDataDriven(int num1, char operation, int num2, int result)
                {
                    _calcStandardView.PerformCalculation(num1, operation, num2);

                    _calcStandardView.AssertResult(Convert.ToDecimal(result));
                }
                */
    }
}
