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
        public WindowsElement ThisFolderButton => _driver.FindElementByName("This folder");
        public WindowsElement AllFoldersButton => _driver.FindElementByName("All folders");
        public WindowsElement CancelButton => _driver.FindElementByName("Cancel");
        /// <summary>
        /// View 
        /// </summary>
        public WindowsElement ViewButton => _driver.FindElementByName("View");
        /// <summary>
        /// Current View
        /// </summary>
        public WindowsElement CurrentViewButton => _driver.FindElementByName("Current View");
        public WindowsElement ChangeViewButton => _driver.FindElementByName("Change View");
        public WindowsElement CompactButton => _driver.FindElementByName("Compact");
        public WindowsElement SingleButton => _driver.FindElementByName("Single");
        public WindowsElement PreviewButton => _driver.FindElementByName("Preview");
       /// <summary>
       /// View Messages
       /// </summary>
        public WindowsElement ShowAsConversationsButton => _driver.FindElementByName("Show as Conversations");
        public WindowsElement MessagesButton => _driver.FindElementByName("Messages");
        /// <summary>
        /// QuickFile
        /// </summary>
        public WindowsElement QuickFileButton => _driver.FindElementByName("QuickFile");
        public WindowsElement ColumnRButton => _driver.FindElementByName("Column right");
        public WindowsElement ChooseFolderButton => _driver.FindElementByName("Choose Folder");
        public WindowsElement GotoFolderButton => _driver.FindElementByName("Goto Folder");
        public WindowsElement HistoryButton => _driver.FindElementByName("History");
        public WindowsElement CloseButton => _driver.FindElementByName("Close");

    }
}
