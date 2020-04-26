using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;

namespace oWinAppDriverPageObjects.Views
{
    public partial class OutLookStandardView
    {
        public void AssertDisplayed(WindowsElement element1)
        {
            //Assert.AreEqual(true, element1.Displayed,);
        }
        /*
        public void AssertResult(decimal expectedReslt)
        {
            string strResult = GetCalculatorResultText();
            var actualResult = decimal.Parse(strResult);

            Assert.AreEqual(expectedReslt, actualResult);
        }
        */
    }
}
