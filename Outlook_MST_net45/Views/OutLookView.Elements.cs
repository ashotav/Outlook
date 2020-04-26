﻿using System;
using OpenQA.Selenium.Appium.Windows;

namespace oWinAppDriverPageObjects.Views
{
    public partial class OutLookStandardView
    {
        public WindowsElement FileButton => _driver.FindElementByName("File");
        public WindowsElement FolderButton => _driver.FindElementByName("Folder");
        public WindowsElement NewFolderButton => _driver.FindElementByName("New Folder");
        public WindowsElement NewE_MailButton => _driver.FindElementByName("New Email");
        public WindowsElement HomeButton => _driver.FindElementByName("Home");
        /*
        public WindowsElement ZeroButton => _driver.FindElementByName("Zero");
        public WindowsElement OneButton => _driver.FindElementByName("One");
        public WindowsElement TwoButton => _driver.FindElementByName("Two");
        public WindowsElement ThreeButton => _driver.FindElementByName("Three");
        public WindowsElement FourButton => _driver.FindElementByName("Four");
        public WindowsElement FiveButton => _driver.FindElementByName("Five");
        public WindowsElement SixButton => _driver.FindElementByName("Six");
        public WindowsElement SevenButton => _driver.FindElementByName("Seven");
        public WindowsElement EightButton => _driver.FindElementByName("Eight");
        public WindowsElement NineButton => _driver.FindElementByName("Nine");
        public WindowsElement PlusButton => _driver.FindElementByName("Plus");
        public WindowsElement MinusButton => _driver.FindElementByName("Minus");
        public WindowsElement EqualsButton => _driver.FindElementByName("Equals");
        public WindowsElement DivideButton => _driver.FindElementByName("Divide by");
        public WindowsElement MultiplyByButton => _driver.FindElementByName("Multiply by");
        public WindowsElement ResultsInput => _driver.FindElementByAccessibilityId("CalculatorResults");
        */
    }
}
