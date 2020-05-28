// <copyright file="BingTests.cs" company="Automate The Planet Ltd.">
// Copyright 2017 Automate The Planet Ltd.
// Licensed under the Apache License, Version 2.0 (the "License");
// You may not use this file except in compliance with the License.
// You may obtain a copy of the License at http://www.apache.org/licenses/LICENSE-2.0
// Unless required by applicable law or agreed to in writing, software
// distributed under the License is distributed on an "AS IS" BASIS,
// WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
// See the License for the specific language governing permissions and
// limitations under the License.
// </copyright>
// <author>Anton Angelov</author>
// <site>https://automatetheplanet.com/</site>

using System;
using System.IO;
using System.Reflection;
using System.Threading;
using HuddlePageObjectsAppDesignPattern;
using HuddlePageObjectsAppDesignPattern.Core;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Edge;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;

namespace HuddlePageObjectsElementsStringProperties
{
    [TestClass]
    public class BingTests
    {
        private App _app;

        [TestInitialize]
        public void TestInitialize()
        {
            _app = new App();
        }

        [TestCleanup]
        public void TestCleanup()
        {
            _app.Dispose();
        }

        [TestMethod]
        public void TryFirefoxDriver()
        {
            //var d = new EdgeDriver();
            //using (var driver = new FirefoxDriver())
            //{
            //    driver.Navigate().GoToUrl("https://automatetheplanet.com/multiple-files-page-objects-item-templates/");
            //    var link = driver.FindElement(By.PartialLinkText("TFS Test API"));
            //    var jsToBeExecuted = $"window.scroll(0, {link.Location.Y});";
            //    ((IJavaScriptExecutor)driver).ExecuteScript(jsToBeExecuted);
            //    var wait = new WebDriverWait(driver, TimeSpan.FromMinutes(1));
            //    /*DBG*/
            //    Thread.Sleep(TimeSpan.FromSeconds(8));
            //    var clickableElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.PartialLinkText("TFS Test API")));
            //    clickableElement.Click();
            //}
        }

        [TestMethod]
        public void UseApp_SearchTextInBing_UseElementsDirectly()
        {
            var bingMainPage = _app.GoTo<BingMainPage>();
            /*DBG*/ Thread.Sleep(TimeSpan.FromSeconds(2));
            bingMainPage.Search("Automate The Planet");
            //Thread.Sleep(TimeSpan.FromSeconds(30));
            string str = "916,000 Results";
            System.Threading.Thread.Sleep(TimeSpan.FromSeconds(8));
            bingMainPage.AssertResultsCount(str);
            //bingMainPage.AssertResultsCount("940,000 Results");
            //System.Threading.Thread.Sleep(TimeSpan.FromSeconds(2));
        }
    }
}
