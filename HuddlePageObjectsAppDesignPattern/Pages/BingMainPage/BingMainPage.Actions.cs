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

namespace HuddlePageObjectsAppDesignPattern
{
    public partial class BingMainPage : AssertedNavigatablePage
    {
        public BingMainPage(IWebDriver driver) : base(driver)
        {
        }

        public override string Url => "http://www.bing.com/";
        //public override string Url => "http://www.google.com/";
        //public override string Url => "http://yandex.com/";


        public void Search(string textToType)
        {
            SearchBox.Clear();
            SearchBox.SendKeys(textToType);
            //SearchBox.SendKeys(Keys.Enter);
            //WrappedDriver.SwitchTo().Frame(WrappedDriver.FindElement(By.Id("sb_form")));
            ////1. presence in DOM
            var wait = new WebDriverWait(WrappedDriver, TimeSpan.FromSeconds(8));
            ////wait.Until(ExpectedConditions.ElementExists(By.Id("sb_form_go")));
            //wait.Until(ExpectedConditions.ElementToBeClickable(By.Id("sb_form_go")));
            Actions act = new Actions(WrappedDriver);
            //act.MoveToElement(GoButton).ClickAndHold();
            //act.MoveToElement(WrappedDriver.FindElement(By.Id("sb_form_go"))).SendKeys(Keys.Enter).Perform();
            //.Perform();
            //wait.Until(ExpectedConditions.VisibilityOfAllElementsLocatedBy(By.XPath("sb_form_go")));
            //GoButton.Submit();
            //string js = string.Format("window.scroll(0, {0});", GoButton.Location.Y);
            //((IJavaScriptExecutor)WrappedDriver).ExecuteScript(js);
            //GoButton.Click();
            //SearchYText.Clear();
            //SearchYText.SendKeys(textToType);
            //SearchYB.Click();
            
            SelectElement select = new SelectElement(WrappedDriver.FindElement(By.Id("sb_form_q")));
            select.SelectByText(textToType);
        }
    }
}
