using AventStack.ExtentReports;
using Microsoft.Office.Interop.Outlook;
using OpenQA.Selenium;
using OpenQA.Selenium.Appium.Windows;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.ServiceProcess;
//using System.Text;
using System.Threading;

using Z01_NU_nCore.Tests;

namespace Z01_NU_nCore.Helpers
{
    public class Helper
    {
        public Application outlookApp = null;
        public NameSpace outlookNamespace = null;
        MAPIFolder inboxFolder = null;

        public Helper()
        {
            outlookApp = new Application();
            outlookNamespace = outlookApp.GetNamespace("MAPI");
            inboxFolder = outlookNamespace.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
        }

        public void cleanupFolder(MAPIFolder folder)
        {
            Items messages = folder.Items;
            while (messages.Count > 0)
            {
                foreach (Object _obj in messages)
                {
                    if (_obj is MailItem)
                    {
                        MailItem message = (MailItem)_obj;
                        message.Delete();
                    }

                }
                messages = folder.Items;
            }

        }

        public void cleanupInbox()
        {
            cleanupFolder(inboxFolder);
        }
        /*        public void cleanupSent() {

                }
                */
        //sort by recive date -inbox
        //sort by sent date - Sent Items
        public void copyEmails(string fromFolderName, string toFolderName, int mailCount)
        {
            MAPIFolder fromFolder = getfolder(fromFolderName);
            MAPIFolder toFolder = getfolder(toFolderName);
            Items fromMessages = fromFolder.Items;
            for (int i = 1; i <= mailCount; i++)
            {
                MailItem message = fromMessages[i];
                MailItem copiedmessage = message.Copy();
                copiedmessage.Move(toFolder);
            }
        }

        public void prms(MailItem mail)
        {
            for (int i = 0; i < 60; i++)
            {
                Console.WriteLine(mail.ItemProperties[i].Name.ToString());
            }
        }
        public Dictionary<string, dynamic> getAllProperiesOfEmail(MailItem mail)
        {
            Dictionary<string, dynamic> resultDictionary = new Dictionary<string, dynamic>();
            string ZeroFilingStatus = (string)mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/ZeroFilingStatus/0x0000001F");
            string Destination = (string)mail.PropertyAccessor.GetProperty("http://schemas.microsoft.com/mapi/string/{00020329-0000-0000-C000-000000000046}/Destination/0x0000001F");
            if (!string.IsNullOrEmpty(ZeroFilingStatus))
            {
                resultDictionary.Add("ZeroFilingStatus", ZeroFilingStatus);
            }
            if (!string.IsNullOrEmpty(Destination))
            {
                resultDictionary.Add("Destination", Destination);
            }
            foreach (ItemProperty property in mail.ItemProperties)
            {
                if (property.Name.ToString() == "Categories" || property.Name.ToString() == "Subject" || property.Name.ToString() == "Attachments")
                {
                    resultDictionary.Add(property.Name.ToString(), property.Value);
                }
            }
            return resultDictionary;
        }

        public dynamic getProperty(MailItem mail, string property)
        {
            return mail.ItemProperties[property].Value;
        }

        public Dictionary<string, dynamic> getAllProperiesOfSelectedEmail()
        {
            return getAllProperiesOfEmail(GetSelectedEmails().First());
        }

        public IEnumerable<MailItem> GetSelectedEmails()
        {
            foreach (MailItem email in outlookApp.ActiveExplorer().Selection)
            {
                yield return email;
            }
        }

        public MAPIFolder getfolder(String folderName, MAPIFolder parentFolder = null)
        {

            if (parentFolder != null && parentFolder.Name == folderName)
                return parentFolder;

            Folders oFolders;

            if (parentFolder == null)
                parentFolder = inboxFolder.Parent;

            oFolders = parentFolder.Folders;
            foreach (MAPIFolder Folder in oFolders)
            {
                string currentFolderName = Folder.Name;
                var found = getfolder(folderName, Folder);
                if (found != null)
                {
                    return found;
                }
            }
            return null;
        }

        public void killZeroProcess()
        {
            ServiceController sc = new ServiceController();
            sc.ServiceName = "ZeroOutlookWcfServicesHost";
            if (sc.Status == ServiceControllerStatus.Running)
            {
                sc.Stop();
                sc.WaitForStatus(ServiceControllerStatus.Stopped);
            }
        }

        public void startZeroProcess()
        {
            ServiceController sc = new ServiceController();
            sc.ServiceName = "ZeroOutlookWcfServicesHost";
            if (sc.Status != ServiceControllerStatus.Running)
            {
                sc.Start();
                sc.WaitForStatus(ServiceControllerStatus.Running);
            }

        }

        public Boolean isAppRuning(WindowsDriver<WindowsElement> driver)
        {
            try
            {
                string tmp = driver.CurrentWindowHandle;
                return true;
            }
            catch
            {
                return false;
            }
        }

        public void waitUntilOutlookOpens(WindowsDriver<WindowsElement> driver)
        {
            BaseTest.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.Name("File Tab")));
        }


        public Boolean isElementExists(IWebElement element)
        {
            try
            {
                return element.Displayed;
            }
            catch
            {
                return false;
            }

        }

        public Boolean isElementExists(WindowsDriver<WindowsElement> driver, By by)
        {
            try
            {
                WindowsElement elem = driver.FindElement(by);
                return true;
            }
            catch
            {
                return false;
            }

        }

        public Boolean isElementExists(IWebElement driver, By by)
        {
            try
            {
                IWebElement elem = driver.FindElement(by);
                return true;
            }
            catch
            {
                return false;
            }

        }

        public void errorHendler(WindowsDriver<WindowsElement> driver, string text)
        {
            Screenshot file = null;
            try
            {
                file = driver.GetScreenshot();
            }
            catch
            {
                try
                {
                    file = driver.GetScreenshot();
                }
                catch { }
            }

            //To save screenshot
            string pngfileName = Guid.NewGuid().ToString();
            //   string pngfilepath = Path.GetDirectoryName(BaseTest.reportPath)+"\\Screenshots\\" + pngfileName + ".png";
            string path = Path.GetDirectoryName(BaseTest.reportPath) + "\\Screenshots";
            if (!Directory.Exists(path))
            {
                Directory.CreateDirectory(path);
            }
            string pngfilepath = path + "\\" + pngfileName + ".png";
            if (file != null)
            {
                file.SaveAsFile(pngfilepath, ScreenshotImageFormat.Png);
                BaseTest.extentTest.Fail(text, MediaEntityBuilder.CreateScreenCaptureFromPath(pngfilepath).Build());
            }
            else
            {
                BaseTest.extentTest.Fail(text + "<br>" + "was not able to add screenshot");
            }
            throw new System.Exception(text);
        }

        public void warningHendler(string text)
        {
            BaseTest.extentTest.Warning(text);
        }

        public void ZipDirectory(String path)
        {
            string parent = Path.GetDirectoryName(path);
            string name = Path.GetFileName(path);
            string filename = Path.Combine(parent, name + ".zip");
            string sipfile = path + ".zip";
            if (File.Exists(sipfile))
            {
                File.Delete(sipfile);
            }
            ZipFile.CreateFromDirectory(path, filename, CompressionLevel.Fastest, true);
        }

        public void waitUntillPageLoad(WindowsDriver<WindowsElement> driver, By by)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
        }

        public void waitUntillPageLoad(By by)
        {
            //  var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            try
            {
                BaseTest.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(by));
            }
            catch (System.Exception)
            {

                errorHendler(BaseTest.driver, "cannot find element in page");
            }

        }
        public void waitUntillElementDoesNotExists(WindowsDriver<WindowsElement> driver, By by)
        {
            var wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
            //    BaseTest.wait.Until(driver => !driver.FindElement(by).Displayed);
            wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(by));

        }

        public void waitUntillElementDoesNotExists(By by)
        {
            BaseTest.wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(by));

        }

        public void changeWindowLevel(WindowsDriver<WindowsElement> driver, int level)
        {

            var allWindowHandles = driver.WindowHandles;
            driver.SwitchTo().Window(allWindowHandles[level]);
        }

        public void setCurrentWindowLevel(WindowsDriver<WindowsElement> driver)
        {
            Thread.Sleep(3000);
            changeWindowLevel(driver, 0);
        }

        public void closeAllApplevels(WindowsDriver<WindowsElement> driver)
        {
            var allWindowHandles = driver.WindowHandles;
            foreach (var item in allWindowHandles)
            {
                driver.SwitchTo().Window(item);
                driver.CloseApp();
            }

        }

        public IWebElement getEmailById(WindowsDriver<WindowsElement> driver, int number)
        {
            number = number + 4;
            string xpath = "//[@Name=\"Table View\"]";
            return getElementById(driver, xpath, number);
        }

        public IWebElement getElementById(WindowsDriver<WindowsElement> driver, string Xpath, int id)
        {
            var elements = driver.FindElements(By.XPath(Xpath + "/*"));
            return elements[id];
        }

    }
}
