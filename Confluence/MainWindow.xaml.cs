using System;
using System.IO;
using System.Windows;
using Confluence.Properties;
using log4net;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;

namespace Confluence
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private FirefoxDriver _driver;
        private const string Notes = @"Note: This page was imported from share point." +
            "The original document had been attached as an attachment.\r\n";

        private ILog _logger = LogManager.GetLogger(typeof(MainWindow));

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

        public abstract class FileImporter
        {
            private FirefoxDriver _driver;

            protected FileImporter(FirefoxDriver driver)
            {
                _driver = driver;
            }

            public void Import(FirefoxDriver driver, FileInfo file)
            {
                CreateParentPages(file.DirectoryName);
            }

            private void CreateParentPages(string filePath)
            {
                var parentPage = GetSpaceRootUrl();
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
                GotoPage(pageName);
                bool exists = false;
                try
                {
                    IWebElement pageTitle = null;
                    PerformActionWithRetry(() => { pageTitle = _driver.FindElementById("title-text"); });
                    exists = pageTitle.Text != "Page Not Found";
                }
                catch
                {
                    // ignored
                }
                return exists;
            }

            private void GotoPage(string pageName)
            {
                var pagePath = GetSpaceRootUrl() + "/" + pageName;
                _driver.Navigate().GoToUrl(pagePath);
            }
            private void CreatePage(string pageName, string pageContent)
            {
                PerformActionWithRetry(() =>
                {
                    var createPageBtn = _driver.FindElement(By.Id("quick-create-page-button"));
                    createPageBtn.Click();
                });


                _driver.FindElement(By.Id("content-title")).SendKeys(pageName);
                _driver.SwitchTo().Frame("wysiwygTextarea_ifr");
                _driver.FindElement(By.Id("tinymce")).SendKeys(pageContent);
                _driver.SwitchTo().ParentFrame();
                SavePage();
            }
            private void SavePage()
            {
                _driver.FindElement(By.Id("rte-button-publish")).Click();
            }
            protected abstract void DoImport();
        }

        private void ButtonBase_OnClick(object sender, RoutedEventArgs e)
        {
            DirectoryInfo dirInfo = new DirectoryInfo(Settings.Default.ImportFrom);
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
                if (DoesPageExist(pageName))
                {
                    pageName = string.Format("Conflict page {0} {1}", pageName, Guid.NewGuid().ToString());
                    _logger.InfoFormat("Page for {0} already exists, create as {1}", 
                        file.Name, pageName);
                }

                GotoSpaceRootPage();
                var fileExt = file.Extension.ToLowerInvariant();
                if (fileExt == ".doc" || fileExt == ".docx")
                {
                    CreateParentPages(file.DirectoryName);
                    ImportWordDoc(file.FullName);
                    File.Move(file.FullName, Path.Combine(file.DirectoryName, file.Name + ".migrated"));
                }
                else if (fileExt == ".xls" || fileExt == ".xlsx")
                {
                    CreateParentPages(file.DirectoryName);
                    CreatePage(pageName, Notes);
                    AttachTheOriginalFile(file.FullName);
                    AddMacroForAttachment(MacroType.Excel, file.Name);
                    File.Move(file.FullName, Path.Combine(file.DirectoryName, file.Name + ".migrated"));
                }
                else if (fileExt == ".pdf")
                {
                    CreateParentPages(file.DirectoryName);
                    CreatePage(pageName, Notes);
                    AttachTheOriginalFile(file.FullName);
                    AddMacroForAttachment(MacroType.Pdf, file.Name);
                    File.Move(file.FullName, Path.Combine(file.DirectoryName, file.Name + ".migrated"));
                }
                else if (fileExt == ".jpg" || fileExt == ".jpeg" || fileExt == ".png")
                {
                    
                }
            }

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
            textarea.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.End);
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
            _driver.FindElementByCssSelector(".button-panel-button.ok").Click(); 
            _driver.FindElementById("rte-button-publish").Click();
        }

        private void ImportWordDoc(string wordDocPath)
        {
            NavigateToImportWordDocPage();
            DoImportWordDoc(wordDocPath);
            AttachTheOriginalFile(wordDocPath);
            AddNotesToPageToSayTheFileWasImportedFromSharepoint();
        }

        private void AddNotesToPageToSayTheFileWasImportedFromSharepoint()
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
            textarea.SendKeys(OpenQA.Selenium.Keys.Control + OpenQA.Selenium.Keys.Home);

            textarea.SendKeys(Notes);
            _driver.SwitchTo().ParentFrame();
            SavePage();
        }

        private void AttachTheOriginalFile(string filePath)
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

        private void DoImportWordDoc(string wordDocPath)
        {
            PerformActionWithRetry(() => { _driver.FindElement(By.Id("filename")).SendKeys(wordDocPath); });
            _driver.FindElement(By.Id("next")).Click();

            PerformActionWithRetry(() => { _driver.FindElement(By.Id("importwordform")).Submit(); });
        }

        

        private void NavigateToImportWordDocPage()
        {
            OpenActionMenuOfPage();
            _driver.FindElement(By.Id("import-word-doc")).Click();
        }

        private void OpenActionMenuOfPage()
        {
            PerformActionWithRetry(() => { _driver.FindElement(By.Id("action-menu-link")).Click(); });
        }

        private void GotoSpaceRootPage()
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

        private void Login()
        {
            var loginUrl = Settings.Default.BaseAddress + (Settings.Default.IsTesting ? "/login.action" : "login");

            _driver.Navigate().GoToUrl(loginUrl);

            PerformActionWithRetry(() =>
            {
                var usernameId = Settings.Default.IsTesting ? "os_username" : "username";
                _driver.FindElement(By.Id(usernameId)).SendKeys(Properties.Settings.Default.UserName);
            });

            var passwordId = Settings.Default.IsTesting ? "os_password" : "password";
            _driver.FindElement(By.Id(passwordId)).SendKeys(Properties.Settings.Default.Password);

            var loginBtnId = Settings.Default.IsTesting ? "loginButton" : "login";
            _driver.FindElement(By.Id(loginBtnId)).Click();
        }
    }
}
