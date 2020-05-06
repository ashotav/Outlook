using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;

namespace WinAppDriverPgeObjects.Views
{
    public partial class OutLookStandardView
    {
        public OutLookPageObjectsTests cont = new OutLookPageObjectsTests();
        private readonly WindowsDriver<WindowsElement> _driver;
        //OutLookPageObjectsTests pObj = new OutLookPageObjectsTests();
        public OutLookStandardView(WindowsDriver<WindowsElement> driver) => _driver = driver;
        //}
        public void WaitUntil(WindowsElement element, WebDriverWait wait, string funcName)
        {
            string txt = element.Text;
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            Console.WriteLine($"Console: Test:{funcName} {txt}{fName} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            wait.Until(pred => element.Displayed);
            Console.WriteLine($"Console: Test:{funcName} {txt} {fName} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
        }

        public void ClickElement(WindowsElement element, string funcName)
        {
            string txt = element.Text;
            Console.WriteLine($"Console: Test:{funcName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            element.Click();
            Console.WriteLine($"Console before scrShot: Test:{funcName} { element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            //pObj.TestContext.WriteLine($"Console before scrShot: Test:{funcName} { element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            var screenShot = _driver.GetScreenshot();
            //screenShot.SaveAsFile($"..\\ScreenShot\\{funcName}{DateTime.Now.ToString("ddMMyyyyhhmmss")}.png", OpenQA.Selenium.ScreenshotImageFormat.Png);
            screenShot.SaveAsFile($@"{OutLookPageObjectsTests.deployDir}\\{funcName}{DateTime.Now.ToString("ddMMyyyyhhmmss")}.png", OpenQA.Selenium.ScreenshotImageFormat.Png);
            Console.WriteLine($"Console after scrShot: Test:{funcName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            cont.TestContext.WriteLine($"Console after scrShot: Test:{funcName} {txt} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");

        }
    }
}
