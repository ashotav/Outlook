//using NUnit.Framework;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Appium.Windows;
using System;
using System.Diagnostics;

namespace WinAppDriverPgeObjects.Views
{
    public partial class OutLookStandardView
    {
        //public void AssertDisplayed(WindowsElement element1) => Assert.IsTrue(element1.Displayed);
        public void AssertDisplayed(WindowsElement element)
        {
            Assert.IsTrue(element.Displayed, $"Not Found {element.Text }");
            Console.WriteLine($"Console Displayed element text: {element.Text} {element.Displayed}");
        }
        public void AssertExpectedText(string element,string expected)
        {
           
            Assert.AreEqual(expected,element, false, $"Not Expected {element }");
        }

    }
}
