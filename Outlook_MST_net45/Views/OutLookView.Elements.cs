using OpenQA.Selenium.Appium.Windows;

namespace WinAppDriverPgeObjects.Views
{
    public partial class OutLookStandardView
    {
        public WindowsElement FileButton => _driver.FindElementByName("File");
        public WindowsElement FolderButton => _driver.FindElementByName("Folder");
        public WindowsElement NewFolderButton => _driver.FindElementByName("New Folder...");
        public WindowsElement NewEmailButton => _driver.FindElementByName("New Email");
        public WindowsElement HomeButton => _driver.FindElementByName("Home");
        public WindowsElement ViewButton => _driver.FindElementByName("View");
        public WindowsElement ChangeViewButton =>  _driver.FindElementByName("Change View");
        public WindowsElement CompactButton => _driver.FindElementByName("IMAP Messages");
        public WindowsElement SingleButton => _driver.FindElementByName("Hide Messages Marked for Deletion");
        public WindowsElement PreviewButton => _driver.FindElementByName("Preview");
        public WindowsElement ThisFolderButton => _driver.FindElementByName("This folder");
        public WindowsElement AllFoldersButton => _driver.FindElementByName("All folders");
        public WindowsElement CancelButton => _driver.FindElementByName("Cancel");
        public WindowsElement ShowAsConversationsButton => _driver.FindElementByName("Show as Conversations");

    }
}
