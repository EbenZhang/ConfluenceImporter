using System;
using System.Diagnostics;
using System.IO;
using System.Threading;
using System.Windows;
using Confluence.Properties;
using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

namespace Confluence
{
    public partial class MainWindow 
    {
        private FirefoxDriver _driver;
        private const string Notes = @"Note: This page was imported from share point." +
            "The original document had been attached as an attachment.\r\n";

        private readonly ILog _logger = LogManager.GetLogger(typeof(MainWindow));

        public MainWindow()
        {
            InitializeComponent();
        }
        static void PerformActionWithRetry(Action action)
        {
            Exception exception = null;
            for (var i = 0; i < 10; ++i)
            {
                try
                {
                    action();
                    return;
                }
                catch (Exception ex)
                {
                    System.Threading.Thread.Sleep(1000);
                    exception = ex;
                    // ignored
                }
            }
            if (exception != null) throw exception;
        }

        public abstract class AbstractFileImporter
        {
            protected FirefoxDriver _driver;

            protected AbstractFileImporter(FirefoxDriver driver)
            {
                _driver = driver;
            }

            public void Import(FileInfo file, string asPage)
            {
                CreateParentPages(file.DirectoryName);
                DoImport(file, asPage);
                Debug.Assert(file.DirectoryName != null, "file.DirectoryName != null");
                file.MoveTo(Path.Combine(file.DirectoryName, file.Name + ".migrated"));
            }

            private void CreateParentPages(string filePath)
            {
                var parentPage = "";// empty for space root page
                var directories = filePath.Replace(Settings.Default.ImportFrom, "")
                    .Split(new[] { Path.DirectorySeparatorChar }, StringSplitOptions.RemoveEmptyEntries);
                foreach (var dir in directories)
                {
                    var exists = DoesPageExist(dir);
                    if (!exists)
                    {
                        GotoPage(parentPage);

                        CreatePage(dir, "");
                    }

                    parentPage = dir;
                }
            }
            private bool DoesPageExist(string pageName)
            {
                return DoesPageExist(_driver, pageName);
            }
            public static bool DoesPageExist(FirefoxDriver driver, string pageName)
            {
                GotoPage(driver, pageName);
                bool exists = false;
                try
                {
                    IWebElement pageTitle = null;
                    PerformActionWithRetry(() => { pageTitle = driver.FindElementById("title-text"); });
                    // Space Tools is the page to recover a deleted page.
                    exists = pageTitle.Text != "Page Not Found" && pageTitle.Text != "Space Tools";
                    if (exists)
                    {
                        // in case the page is in another space. 
                        // confluence shows the title with the correct page name
                        // but the page does not exist.
                        exists = driver.FindElementByCssSelector("div#content div.aui-message p.title").Text != "Page Not Found";
                    }
                }
                catch
                {
                    // ignored
                }
                return exists;
            }

            private void GotoPage(string pageName)
            {
                GotoPage(_driver, pageName);
            }
            private static void GotoPage(FirefoxDriver driver, string pageName)
            {
                var pagePath = GetSpaceRootUrl() + "/" + pageName;
                driver.Navigate().GoToUrl(pagePath);
            }
            protected void CreatePage(string pageName, string pageContent)
            {
                PerformActionWithRetry(() =>
                {
                    var createPageBtn = _driver.FindElement(By.Id("quick-create-page-button"));
                    createPageBtn.Click();
                });

                PerformActionWithRetry(() =>
                {
                    _driver.FindElement(By.Id("content-title")).SendKeys(pageName);
                });
                _driver.SwitchTo().Frame("wysiwygTextarea_ifr");
                _driver.FindElement(By.Id("tinymce")).SendKeys(pageContent);
                _driver.SwitchTo().ParentFrame();
                SavePage();
            }
            
            private void SavePage()
            {
                _driver.FindElement(By.Id("rte-button-publish")).Click();
            }

            protected void AddNotesToPageToSayTheFileWasImportedFromSharepoint()
            {
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementById("editPageLink").Click();
                });

                PerformActionWithRetry(() =>
                {
                    _driver.SwitchTo().Frame("wysiwygTextarea_ifr");
                });

                var textarea = _driver.FindElementById("tinymce");
                textarea.Click();
                textarea.SendKeys(Keys.Control + Keys.Home);

                textarea.SendKeys(Notes);
                _driver.SwitchTo().ParentFrame();
                SavePage();
            }
            protected void AttachTheOriginalFile(string filePath)
            {
                NavigateToUploadAttachmentPage();
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementById("file_0").SendKeys(filePath);
                });

                _driver.FindElementById("upload-attachments").Submit();
                _driver.FindElementById("viewPageLink").Click();
            }
            private void NavigateToUploadAttachmentPage()
            {
                OpenActionMenuOfPage();
                _driver.FindElement(By.Id("view-attachments-link")).Click();
            }

            protected void NavigateToImportWordDocPage()
            {
                OpenActionMenuOfPage();
                _driver.FindElement(By.Id("import-word-doc")).Click();
            }

            private void OpenActionMenuOfPage()
            {
                PerformActionWithRetry(() => { _driver.FindElement(By.Id("action-menu-link")).Click(); });
            }

            public static void GotoSpaceRootPage(FirefoxDriver _driver)
            {
                var spaceRoot = GetSpaceRootUrl();
                _driver.Navigate().GoToUrl(spaceRoot);
            }

            private static string GetSpaceRootUrl()
            {
                var wikiUrl = Settings.Default.WikiBaseUrl;
                var spaceRoot = wikiUrl + "/display/" + Settings.Default.SpaceName;
                return spaceRoot;
            }

            protected abstract void DoImport(FileInfo file, string asPage);

            protected void PublishPage()
            {
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementById("rte-button-publish").Click();
                    IWebElement editorIFrame = null;
                    try
                    {
                        editorIFrame = _driver.FindElementById("wysiwygTextarea_ifr");
                    }
                    catch
                    {// the page saved sucessfully if cannot find the iframe
                        return;
                    }
                    if (editorIFrame != null)
                    {
                        throw new Exception("Throw exception to retry");
                    }
                });
            }
        }

        class WordDocImporter : AbstractFileImporter
        {
            public WordDocImporter(FirefoxDriver driver) : base(driver)
            {
            }

            protected override void DoImport(FileInfo file, string asPage)
            {
                ImportWordDoc(file.FullName, asPage);
            }

            private void ImportWordDoc(string wordDocPath, string asPage)
            {
                CreatePage(asPage, "");
                NavigateToImportWordDocPage();
                DoImportWordDoc(wordDocPath, asPage);
                AttachTheOriginalFile(wordDocPath);
                AddNotesToPageToSayTheFileWasImportedFromSharepoint();
            }
            private void DoImportWordDoc(string wordDocPath, string pageName)
            {
                PerformActionWithRetry(() => { _driver.FindElement(By.Id("filename")).SendKeys(wordDocPath); });
                _driver.FindElement(By.Id("next")).Click();
                PerformActionWithRetry(() => {
                    _driver.FindElementById("overwritepage").Click();
                });
                _driver.FindElement(By.Id("importwordform")).Submit();
            }
        }

        class MacroableFileImporter : AbstractFileImporter
        {
            public MacroableFileImporter(FirefoxDriver driver) : base(driver)
            {
            }

            protected override void DoImport(FileInfo file, string asPage)
            {
                CreatePage(asPage, Notes);
                AttachTheOriginalFile(file.FullName);
                AddMacroForAttachment(GetMacroType(file.Extension.ToLowerInvariant()), file.Name);
            }

            private static MacroType GetMacroType(string fileExt)
            {
                if(fileExt == ".xls" || fileExt == ".xlsx")
                {
                    return MacroType.Excel;
                }
                return MacroType.Pdf;
            }

            private enum MacroType
            {
                Pdf,
                Excel
            };

            private void AddMacroForAttachment(MacroType macro, string fileName)
            {
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementById("editPageLink").Click();
                });

                PerformActionWithRetry(() =>
                {
                    _driver.SwitchTo().Frame("wysiwygTextarea_ifr");
                });

                var textarea = _driver.FindElementById("tinymce");
                textarea.Click();
                textarea.SendKeys(Keys.Control + Keys.End);
                _driver.SwitchTo().ParentFrame();

                _driver.FindElementById("rte-button-insert").Click();

                _driver.FindElementById("rte-insert-macro").Click();

                _driver.FindElementById("macro-browser-search").SendKeys(macro.ToString());
                if (macro == MacroType.Pdf)
                {
                    _driver.FindElementById("macro-viewpdf").Click();
                }
                else
                {
                    _driver.FindElementById("macro-viewxls").Click();
                }
                PerformActionWithRetry(() =>
                {
                    if(_driver.FindElementById("macro-param-name").Text == "No appropriate attachments")
                    {
                        throw new Exception("Throw an exception to try again");
                    }
                });
                _driver.FindElementByCssSelector(".button-panel-button.ok").Click();

                _driver.SwitchTo().Frame("wysiwygTextarea_ifr");

                PerformActionWithRetry(() =>
                {
                    _driver.FindElementByClassName("editor-inline-macro");
                });
                _driver.SwitchTo().ParentFrame();
                PublishPage();
            }
        }

        class ImageImporter : AbstractFileImporter
        {
            public ImageImporter(FirefoxDriver driver) : base(driver)
            {
            }

            protected override void DoImport(FileInfo file, string asPage)
            {
                CreatePage(asPage, Notes);
                AttachTheOriginalFile(file.FullName);
                ShowImageInPage();
            }

            private void ShowImageInPage()
            {
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementById("editPageLink").Click();
                });
                PerformActionWithRetry(() =>
                {
                    _driver.SwitchTo().Frame("wysiwygTextarea_ifr");
                    _driver.FindElement(By.Id("tinymce")).SendKeys(Keys.Control + Keys.End);
                    _driver.SwitchTo().ParentFrame();
                });
                _driver.FindElementByCssSelector("#confluence-insert-files a.toolbar-trigger.aui-button")
                    .Click();
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementsByCssSelector("#attached-files ul.file-list li.attached-file")[0].Click();
                });
                _driver.FindElementByCssSelector(".button-panel-button.insert").Click();
                _driver.SwitchTo().Frame("wysiwygTextarea_ifr");
                PerformActionWithRetry(() =>
                {
                    _driver.FindElementByClassName("confluence-embedded-image");
                });
                _driver.SwitchTo().ParentFrame();
                PublishPage();
            }
        }

        class RegularFileImporter : AbstractFileImporter
        {
            public RegularFileImporter(FirefoxDriver driver) : base(driver)
            {
            }

            protected override void DoImport(FileInfo file, string asPage)
            {
                CreatePage(asPage, Notes);
                AttachTheOriginalFile(file.FullName); 
            }
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            var dirInfo = new DirectoryInfo(Settings.Default.ImportFrom);
            if (!dirInfo.Exists)
            {
                MessageBox.Show(this, string.Format("Folder {0} does not exist", Settings.Default.ImportFrom));
                return;
            }

            _driver = new FirefoxDriver();

            Login();

            foreach (var file in dirInfo.EnumerateFiles("*.*", SearchOption.AllDirectories))
            {
                if (file.Extension == ".migrated")
                {
                    continue;
                }
                var pageName = Path.GetFileNameWithoutExtension(file.Name);
                if (AbstractFileImporter.DoesPageExist(_driver, pageName))
                {
                    pageName = string.Format("Conflict page {0} {1}", file.Name, Guid.NewGuid().ToString());
                    _logger.InfoFormat("Page for {0} already exists, create as {1}", 
                        file.Name, pageName);
                }

                AbstractFileImporter.GotoSpaceRootPage(_driver);

                var importer = GetImporter(file);
                if(importer != null)
                {
                    importer.Import(file, pageName);
                }
            }
        }

        private AbstractFileImporter GetImporter(FileSystemInfo file)
        {
            AbstractFileImporter importer = null;
            var fileExt = file.Extension.ToLowerInvariant();
            switch (fileExt)
            {
                case ".doc":
                case ".docx":
                    importer = new WordDocImporter(_driver);
                    break;
                case ".xls":
                case ".xlsx":
                case ".pdf":
                    importer = new MacroableFileImporter(_driver);
                    break;
                case ".jpg":
                case ".jpeg":
                case ".png":
                    importer = new ImageImporter(_driver);
                    break;
                case ".vsd":
                case ".vsdx":
                    importer = new RegularFileImporter(_driver);
                    break;
            }
            return importer;
        }

        private void Login()
        {
            var loginUrl = Settings.Default.BaseAddress + (Settings.Default.IsTesting ? "/login.action" : "/login");

            _driver.Navigate().GoToUrl(loginUrl);

            PerformActionWithRetry(() =>
            {
                var usernameId = Settings.Default.IsTesting ? "os_username" : "username";
                _driver.FindElement(By.Id(usernameId)).SendKeys(Settings.Default.UserName);
            });

            var passwordId = Settings.Default.IsTesting ? "os_password" : "password";
            _driver.FindElement(By.Id(passwordId)).SendKeys(Settings.Default.Password);

            var loginBtnId = Settings.Default.IsTesting ? "loginButton" : "login";
            _driver.FindElement(By.Id(loginBtnId)).Click();
        }
    }
}
