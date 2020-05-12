// <copyright file="BingMainPage.Actions.cs" company="Automate The Planet Ltd.">
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
using HuddlePageObjectsAppDesignPattern.Core;
using OpenQA.Selenium;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using System;
using System.Linq;

namespace HuddlePageObjectsAppDesignPattern
{
    public partial class BingMainPage : AssertedNavigatablePage
    {
        public BingMainPage(IWebDriver driver) : base(driver)
        {
        }

        public override string Url => "http://www.bing.com/";
        //public override string Url => "http://www.google.com/";
        //public override string Url => "http://yahoo.com/";
        //public override string Url => "http://aol.com/";


        public void Search(string textToType)
        {
            SearchBox.Clear();
            SearchBox.SendKeys(textToType +Keys.Enter);//
            //GoButton.Click();
            ////var link = WrappedDriver.FindElement(By.PartialLinkText("Homepage"));
            //////var jsToBeExecuted = $"window.scroll(0, {link.Location.Y});";
            //////((IJavaScriptExecutor)WrappedDriver).ExecuteScript(jsToBeExecuted);
            ////var wait = new WebDriverWait(WrappedDriver, TimeSpan.FromMinutes(1));
            //////System.Threading.Thread.Sleep(TimeSpan.FromSeconds(8));
            ////var clickableElement = wait.Until(ExpectedConditions.ElementToBeClickable(By.PartialLinkText("Homepage")));
            ////clickableElement.Click();
            //System.Threading.Thread.Sleep(TimeSpan.FromSeconds(8));
            WrappedDriver.SwitchTo().Window(WrappedDriver.WindowHandles.Last());

            WebDriverWait wait = new WebDriverWait(WrappedDriver, TimeSpan.FromSeconds(10));
            wait.Until(ExpectedConditions.PresenceOfAllElementsLocatedBy(By.Id("b_tween")));
            bool bis = ResultsCountDiv.Text.Contains(ResultsCountDiv.Text);
            Assert.IsTrue(bis);
        }
    }
}
