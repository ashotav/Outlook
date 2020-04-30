using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Diagnostics;

namespace WinAppDriverPgeObjects.Views
{
    public partial class OutLookStandardView
    {
        private readonly WindowsDriver<WindowsElement> _driver;

        public OutLookStandardView(WindowsDriver<WindowsElement> driver) => _driver = driver;
        //}
        public void WaitUntil(WindowsElement element, WebDriverWait wait, string funcName)
        {
            string fName = System.Reflection.MethodBase.GetCurrentMethod().Name;
            Debug.WriteLine($"Debug: Test: {funcName} {element.Text} {fName} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            Console.WriteLine($"Console: Test:{funcName} { element.Text}{fName} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            wait.Until(pred => element.Displayed);
            Debug.WriteLine($"Debug: Test: {funcName} {element.Text} {fName} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            Console.WriteLine($"Console: Test:{funcName} { element.Text} {fName} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
        }

        public void ClickElement(WindowsElement element,string funcName)
        {
            Debug.WriteLine($"Debug: Test: {funcName} {element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            Console.WriteLine($"Console: Test:{funcName} { element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            element.Click();
            Debug.WriteLine($"Debug before scrShot: Test: {funcName} {element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            Console.WriteLine($"Console before scrShot: Test:{funcName} { element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            var screenShot = _driver.GetScreenshot();
            screenShot.SaveAsFile($"..\\ScreenShot\\{funcName}{DateTime.Now.ToString("ddMMyyyyhhmmss")}.png", OpenQA.Selenium.ScreenshotImageFormat.Png);
            Debug.WriteLine($"Debug after scrShot: Test: {funcName} {element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
            Console.WriteLine($"Console after scrShot: Test:{funcName} { element.Text} Date: {DateTime.Now:dd-MM-yyyy HH:mm:ss.fff}");
        }
    }
}
