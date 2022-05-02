namespace Automation.Core.Web
{
    using System;
    using System.Collections.Generic;
    using System.Collections.Immutable;
    using System.Collections.ObjectModel;
    using System.Linq;
    using System.Linq.Expressions;
    using System.Reflection;
    using System.Threading;
    using Argument.Check;
    using Newtonsoft.Json;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Support.UI;
    using Serilog;

    /// <summary>
    /// Base page for all Page Objects to extend to when automating Web UI.
    /// </summary>
    public class BasePage
    {
        private readonly IWebDriver _webDriver;

        /// <summary>
        /// Initializes a new instance of the <see cref="BasePage"/> class.
        /// </summary>
        /// <param name="webDriver">IWebDriver instance.</param>
        /// <param name="timeout">Default timeout.</param>
        public BasePage(IWebDriver webDriver, int timeout)
        {
            _webDriver = webDriver;
            Timeout = timeout;
        }

        private int Timeout { get; set; }

        /// <summary>
        /// Fetches file content of locator details from File mentioned
        /// in LocatorFileAttribute which is a class level attribute. Then this
        /// initializes each variable annotated with LocatorAttribute. The string 
        /// mentioned in the locator gets matched with file content. If no such 
        /// locator string found custom exception is thrown which is visible in log.
        /// </summary>
        protected void Initialize()
        {
            string content;
            try
            {
                content = GetFileContent();
                Log.Information("Fetching locators json content successful.");
            }
            catch
            {
                Log.Information("Failed to fetch locators json content.");
                throw;
            }

            ImmutableDictionary<string, object> locatorIdentifiers = FetchLocatorIdentifiers(content);
            ImmutableDictionary<string, KeyValuePair<string, string>> locators = FetchLocators(locatorIdentifiers);
            FieldInfo[] fields = GetType().GetFields(BindingFlags.NonPublic | BindingFlags.Instance);
            Log.Information($"Binding locator fields for page {(string)locatorIdentifiers["Name"]}");

            foreach (FieldInfo field in fields)
            {
                LocatorAttribute attribute = field.GetCustomAttribute<LocatorAttribute>();
                if (attribute != null)
                {
                    try
                    {
                        Log.Information($"Initializing {attribute.Name} with type {locators[attribute.Name].Key} and value {locators[attribute.Name].Value}");
                        UIElement uiElement = new UIElement(_webDriver, locators[attribute.Name].Key, locators[attribute.Name].Value, Timeout);
                        field.SetValue(this, uiElement);
                    }
                    catch (KeyNotFoundException e)
                    {
                        string errorMessage = $"{attribute.Name} is either not present in JSON or there is a mismatch between the " +
                            $"text in JSON and Page Object attribute name {e.Message}";
                        Log.Error(errorMessage);
                        throw new Exception(errorMessage);
                    }
                }
            }
        }

        /// <summary>
        /// Opens Page by default if URL mentioned in Locator file in url key.
        /// Can be used in child classes to open web page explicitly.
        /// </summary>
        /// <param name="url">Page to open.</param>
        protected void OpenWebPage(Uri url)
        {
            _webDriver.Navigate().GoToUrl(url);
            Log.Information($"Opened webpage {url}");
        }

        /// <summary>
        /// Refreshes current page.
        /// </summary>
        protected void RefreshWebPage()
        {
            _webDriver.Navigate().Refresh();
            Log.Information("Refreshed webpage");
        }

        /// <summary>
        /// Fetch current page URL.
        /// </summary>
        /// <returns>URL as string.</returns>
        protected string GetPageURL()
        {
            return _webDriver.Url;
        }

        /// <summary>
        /// Scroll Window.
        /// </summary>
        protected void ScrollWindow()
        {
            Log.Information($"Scrolling window.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("scroll(0, 250);");
        }

        /// <summary>
        /// Wait for auto closure of tab in cases of some downloads.
        /// </summary>
        /// <param name="currentTabCount">Tab count before download click.</param>
        protected void WaitForDownloadTabToClose(int currentTabCount)
        {
            while (_webDriver.WindowHandles.Count > currentTabCount)
            {
                Thread.Sleep(1000);
            }
        }

        /// <summary>
        /// Wait for AJAX calls to complete. Useful for dynamic page loads.
        /// </summary>
        /// <param name="throwException">Force to not throw exception</param>
        protected void WaitForAjax(bool throwException = false)
        {
            for (int i = 0; i < Timeout; i++)
            {
                bool isAjaxComplete = (bool)(_webDriver as IJavaScriptExecutor).ExecuteScript("return jQuery.active == 0");
                if (isAjaxComplete)
                {
                    return;
                }
            }

            if (throwException)
            {
                Log.Error("WebDriver timed out waiting for AJAX call to complete");
                throw new Exception("WebDriver timed out waiting for AJAX call to complete");
            }
        }

        /// <summary>
        /// Wait for any expected page alert.
        /// </summary>
        /// <param name="maxWaitTime">Max time (in seconds) to wait for alert to appear.</param>
        protected void WaitForAlert(int maxWaitTime)
        {
            int waitTime = 0;
            while (waitTime++ < maxWaitTime)
            {
                try
                {
                    Log.Information("Trying to switch to alert.");
                    _webDriver.SwitchTo().Alert();
                    return;
                }
                catch (NoAlertPresentException)
                {
                    Thread.Sleep(1000);
                }
            }
        }

        /// <summary>
        /// Wait for JS to complete loading.
        /// </summary>
        protected void WaitForReady()
        {
            WebDriverWait wait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(Timeout));
            wait.Until(driver =>
            {
                bool isAjaxFinished = (bool)((IJavaScriptExecutor)driver).
                    ExecuteScript("return jQuery.active == 0");
                bool isLoaderHidden = ((IJavaScriptExecutor)driver).
                    ExecuteScript("return document.readyState").Equals("complete");

                return isAjaxFinished & isLoaderHidden;
            });
        }

        /// <summary>
        /// Check if alert is present.
        /// </summary>
        /// <returns>True is alert present else false</returns>
        protected bool IsAlertPresent()
        {
            try
            {
                _webDriver.SwitchTo().Alert();
                Log.Information("Switched to alert.");
                return true;
            }
            catch (NoAlertPresentException)
            {
                Log.Information("No alert to switch to.");
                return false;
            }
        }

        /// <summary>
        /// Accept alert is present.
        /// </summary>
        protected void AcceptAlert()
        {
            _webDriver.SwitchTo().Alert().Accept();
            Log.Information("Alert accepted.");
        }

        /// <summary>
        /// Switch tabs with specified title.
        /// </summary>
        /// <param name="predicateExp">Predicate expression of title comparison.</param>
        protected void SwitchToWindow(Expression<Func<IWebDriver, bool>> predicateExp)
        {
            Throw.IfNull(() => predicateExp);
            Func<IWebDriver, bool> predicate = predicateExp.Compile();
            foreach (string handle in _webDriver.WindowHandles)
            {
                _webDriver.SwitchTo().Window(handle);
                if (predicate(_webDriver))
                {
                    return;
                }
            }

            throw new ArgumentException(string.Format("Unable to find window with condition: '{0}'", predicateExp.Body));
        }

        /// <summary>
        /// Wait for new tab to open.
        /// </summary>
        /// <param name="maxWaitTime">Max time (in seconds) to wait for new tab to open.</param>
        /// <returns>True if new window opened.</returns>
        protected bool HasNewTabOpened(int maxWaitTime)
        {
            int counter = 0;
            while (counter++ < maxWaitTime)
            {
                try
                {
                    ReadOnlyCollection<string> windowHandles = _webDriver.WindowHandles;
                    if (windowHandles.Count > 1)
                    {
                        return true;
                    }

                    Thread.Sleep(1000);
                }
                catch
                {
                    return false;
                }
            }

            return false;
        }

        /// <summary>
        /// Create New Tab.
        /// </summary>
        protected void CreateNewTab()
        {
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("window.open();");
            SwitchToCurrent();
        }

        /// <summary>
        /// Gives the control to newly activated tab.
        /// </summary>
        protected void SwitchToCurrent()
        {
            _webDriver.SwitchTo().Window(_webDriver.WindowHandles.Last());
        }

        /// <summary>
        /// Gives the control to newly activated tab.
        /// </summary>
        protected void SwitchToActiveElement()
        {
            _webDriver.SwitchTo().ActiveElement();
        }

        /// <summary>
        /// Fetches the count of windows open.
        /// </summary>
        /// <returns>Number of windows open.</returns>
        protected int GetCountOfWindowsOpen()
        {
            return _webDriver.WindowHandles.Count;
        }

        /// <summary>
        /// Close all tabs except the primary one.
        /// </summary>
        /// <param name="primaryHandle">Primary tab handle.</param>
        protected void CloseSecondaryTabs(string primaryHandle)
        {
            foreach (string handle in _webDriver.WindowHandles)
            {
                if (!handle.Equals(primaryHandle))
                {
                    _webDriver.SwitchTo().Window(handle);
                    _webDriver.Close();
                }
            }

            _webDriver.SwitchTo().Window(primaryHandle);
        }

        /// <summary>
        /// Close currently active tab.
        /// </summary>
        protected void CloseCurrentTab()
        {
            _webDriver.Close();
            SwitchToCurrent();
        }

        /// <summary>
        /// Quit WebDriver.
        /// </summary>
        protected void Quit()
        {
            _webDriver.Quit();
        }

        /// <summary>
        /// Get the content of file mentioned in LocatorFileAttribute.
        /// </summary>
        /// <returns>File content as string.</returns>
        private string GetFileContent()
        {
            return GetType().GetCustomAttributes(typeof(LocatorResourceAttribute), true).FirstOrDefault()
                 is LocatorResourceAttribute resourceAttribute ? resourceAttribute.Content : null;
        }

        /// <summary>
        /// Convert file content into a dictionary of key and values pair.
        /// </summary>
        /// <param name="content">Pass file content fetched from GetFileContent.</param>
        /// <returns>Immutable Dictionary of content as KeyValuePairs.</returns>
        private ImmutableDictionary<string, object> FetchLocatorIdentifiers(string content)
        {
            LocatorJSONObject locatorJSONObject = JsonConvert.DeserializeObject<LocatorJSONObject>(content);

            return new Dictionary<string, object>
            {
                { "Name", locatorJSONObject.Name },
                { "Page Elements", locatorJSONObject.Elements },
            }.ToImmutableDictionary();
        }

        /// <summary>
        /// Convert each locator JSON object into a KeyValuePair of Locator Name and Locator Details.
        /// </summary>
        /// <param name="locatorIdentifiers">Immutable Dictionary of file content.</param>
        /// <returns>Immutable Dictionary of Locator name as key and KeyValuePair of
        /// locator type and locator value as value.</returns>
        private ImmutableDictionary<string, KeyValuePair<string, string>> FetchLocators(ImmutableDictionary<string, object> locatorIdentifiers)
        {
            Dictionary<string, KeyValuePair<string, string>> locators = new Dictionary<string, KeyValuePair<string, string>>();

            List<PageElement> pageElements = (List<PageElement>)locatorIdentifiers["Page Elements"];
            foreach (PageElement pageElement in pageElements)
            {
                locators.Add(pageElement.Name, new KeyValuePair<string, string>(pageElement.Type, pageElement.Value));
            }

            return locators.ToImmutableDictionary();
        }
    }
}
