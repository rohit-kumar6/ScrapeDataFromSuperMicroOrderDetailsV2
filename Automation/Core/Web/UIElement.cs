namespace Automation.Core.Web
{
    using System;
    using System.Collections.Generic;
    using System.Drawing;
    using System.IO;
    using System.Web;
    using Automation.Utils;
    using OpenQA.Selenium;
    using OpenQA.Selenium.Interactions;
    using OpenQA.Selenium.Support.UI;
    using Serilog;

    /// <summary>
    /// Wrapper class of Selenium WebElement for better UI operations without
    /// the client having to worry about syncs.
    /// </summary>
    public class UIElement
    {
        private readonly IWebDriver _webDriver;
        private readonly WebDriverWait _wait;
        private readonly By _by;
        private readonly int _timeout;

        /// <summary>
        /// Initializes a new instance of the <see cref="UIElement"/> class.
        /// </summary>
        /// <param name="webDriver">WebDriver instance.</param>
        /// <param name="locatorType">Locator type.</param>
        /// <param name="locatorValue">Value of the locator according to type.</param>
        /// <param name="timeout">Default timeout for the element.</param>
        public UIElement(IWebDriver webDriver, string locatorType, string locatorValue, int timeout)
        {
            _webDriver = webDriver;
            LocatorType = locatorType;
            LocatorValue = locatorValue;
            _wait = new WebDriverWait(webDriver, TimeSpan.FromSeconds(timeout));
            _by = GetBy();
            _timeout = timeout;
        }

        /// <summary>
        /// Gets or Sets the type of the locator like id, xpath, css etc...
        /// </summary>
        public string LocatorType { get; set; }

        /// <summary>
        /// Gets or Sets the locator value as per the type.
        /// </summary>
        public string LocatorValue { get; set; }

        /// <summary>
        /// Move to UIElement.
        /// </summary>
        public void MoveToElement()
        {
            Actions actions = new Actions(_webDriver);
            actions.MoveToElement(FindElement(_by));
            actions.Perform();
        }

        /// <summary>
        /// Clears text in the TextBox UIElement.
        /// </summary>
        public void Clear()
        {
            IWebElement webElement = FindElement(_by);
            webElement.Clear();
            Log.Information($"Cleared text in {LocatorValue}.");
        }

        /// <summary>
        /// Sets text in the TextBox UIElement.
        /// </summary>
        /// <param name="value">Value to set in TextBox.</param>
        public void SetText(string value)
        {
            IWebElement webElement = FindElement(_by);
            webElement.Clear();
            webElement.SendKeys(value);
            Log.Information($"Cleared and text set in {LocatorValue} with value: {value}.");
        }

        /// <summary>
        /// JS Sets text in the TextBox UIElement.
        /// </summary>
        /// <param name="value">Value to set in TextBox.</param>
        public void JSSetText(string value)
        {
            Log.Information($"Performing JSSetText on {LocatorValue}.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript($"arguments[0].value=\"{value}\";", FindElement(_by));
        }

        /// <summary>
        /// Sets text without clearing in the TextBox UIElement.
        /// </summary>
        /// <param name="value">Value to set in TextBox.</param>
        public void SetTextWithoutClear(string value)
        {
            IWebElement webElement = FindElement(_by);
            webElement.SendKeys(value);
            Log.Information($"Cleared and text set in {LocatorValue} with value: {value}.");
        }

        /// <summary>
        /// Sets text by selecting and overwriting in the TextBox UIElement.
        /// </summary>
        /// <param name="value">Value to set in TextBox.</param>
        public void SetTextByOverwriting(string value)
        {
            IWebElement webElement = FindElement(_by);
            webElement.SendKeys(Keys.Control + "a");
            webElement.SendKeys(value);
            Log.Information($"Selected all and text set in {LocatorValue} with value: {value}.");
        }

        /// <summary>
        /// Check is UI Element exists.
        /// </summary>
        /// <param name="customTimeOut">Custom Timeout.</param>
        /// <returns>True if UI Element exists else false.</returns>
        public bool Exists(int customTimeOut)
        {
            Log.Information($"Checking existence of {LocatorValue}.");
            WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(customTimeOut));

            try
            {
                customWait.Until(webDriver => webDriver.FindElement(_by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if element is enabled on UI.
        /// </summary>
        /// <returns>True if UI Element is enabled else false.</returns>
        public bool IsEnabled()
        {
            Log.Information($"Checking visibility of {LocatorValue}.");

            return FindElement(_by).Enabled;
        }

        /// <summary>
        /// Check if element is visible on UI.
        /// </summary>
        /// <param name="customTimeOut">Custom Timeout.</param>
        /// <returns>True if UI Element is visible else false.</returns>
        public bool IsVisible(int customTimeOut)
        {
            Log.Information($"Checking visibility of {LocatorValue}.");
            WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(customTimeOut));

            try
            {
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if element is visible on UI.
        /// </summary>
        /// <param name="customTimeOut">Custom Timeout.</param>
        /// <returns>True if UI Element is visible else false.</returns>
        public bool IsClickable(int customTimeOut)
        {
            Log.Information($"Checking if {LocatorValue} is clickable.");
            WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(customTimeOut));

            try
            {
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(_by));
                return true;
            }
            catch (Exception)
            {
                return false;
            }
        }

        /// <summary>
        /// Check if UI Element is selected or not.
        /// </summary>
        /// <returns>True if selected else false.</returns>
        public bool IsSelected()
        {
            Log.Information($"Checking if {LocatorValue} is already selected.");
            return FindElement(_by).Selected;
        }

        /// <summary>
        /// Get inner text of UI Element.
        /// </summary>
        /// <returns>Inner text as string.</returns>
        public string GetText()
        {
            Log.Information($"Getting text of {LocatorValue}.");
            return FindElement(_by).Text;
        }

        /// <summary>
        /// Get the color of the text as a string.
        /// </summary>
        /// <returns>String in RGBA format.</returns>
        public string GetTextColor()
        {
            Log.Information($"Getting text color of {LocatorValue}.");
            return FindElement(_by).GetCssValue("color");
        }

        /// <summary>
        /// Get size of table.
        /// </summary>
        /// <returns>Count of row/columns.</returns>
        public int GetSize()
        {
            Log.Information($"Getting size of {LocatorValue}.");
            var element = FindElement(_by);
            return element.FindElements(By.XPath(LocatorValue)).Count;
        }

        /// <summary>
        /// Get text of parent node only excluding child node texts.
        /// </summary>
        /// <returns>Text of parent node.</returns>
        public string GetTextOfParent()
        {
            Log.Information($"Getting parent text of {LocatorValue}.");
            var element = FindElement(_by);
            var text = element.Text.Trim();
            var children = element.FindElements(By.XPath("./*"));
            foreach (IWebElement child in children)
            {
                text = StringUtils.ReplaceFirst(text, child.Text, string.Empty);
            }

            return text;
        }

        /// <summary>
        /// Click the UIElement.
        /// </summary>
        public void Click()
        {
            Log.Information($"Clicking {LocatorValue}");
            FindElement(_by).Click();
        }

        /// <summary>
        /// Double Click the UIElement.
        /// </summary>
        public void DoubleClick()
        {
            Log.Information($"Right clicking {LocatorValue}");
            var actions = new Actions(_webDriver);
            actions.DoubleClick(FindElement(_by)).Perform();
        }

        /// <summary>
        /// Send Enter Command.
        /// </summary>
        public void SendEnterCommand()
        {
            Log.Information("Sending Enter command.");
            IWebElement webElement = FindElement(_by);
            webElement.SendKeys(Keys.Enter);
        }

        /// <summary>
        /// Get the selected text from the select box.
        /// </summary>
        /// <returns>String selected in select box.</returns>
        public string GetFirstSelectedElement()
        {
            var selectElement = new SelectElement(FindElement(_by));
            return selectElement.SelectedOption.Text;
        }

        /// <summary>
        /// Switch to iFrame.
        /// </summary>
        public void SwitchToIFrame()
        {
            _webDriver.SwitchTo().Frame(FindElement(_by));
            Log.Information($"Switched IFrame to {LocatorValue}");
        }

        /// <summary>
        /// Write in iFrame.
        /// </summary>
        public void WriteInIFrame(string text)
        {
            SwitchToIFrame();
            try
            {
                IWebElement editable = _webDriver.SwitchTo().ActiveElement();
                editable.SendKeys(text);
            }
            catch (Exception ex)
            {
                Log.Error("Exception in iFrame writing: " + ex.Message + ":" + ex.StackTrace);
                throw;
            }
            finally
            {
                SwitchToMainWindowFromIFrame();
            }
        }

        /// <summary>
        /// Read from iFrame.
        /// </summary>
        public string ReadFromIFrame()
        {
            SwitchToIFrame();
            try
            {
                IWebElement editable = _webDriver.SwitchTo().ActiveElement();
                return editable.Text;
            }
            catch (Exception ex)
            {
                Log.Error("Exception in iFrame writing: " + ex.Message + ":" + ex.StackTrace);
                throw;
            }
            finally
            {
                SwitchToMainWindowFromIFrame();
            }
        }

        /// <summary>
        /// Switch to Main Window from iFrame.
        /// </summary>
        public void SwitchToMainWindowFromIFrame()
        {
            _webDriver.SwitchTo().DefaultContent();
        }

        /// <summary>
        /// Wait for UIElement to be clickable and then click.
        /// </summary>
        /// <param name="timeout">Custom timeout.</param>
        public void WaitAndClick(int timeout)
        {
            Log.Information($"Wait for {LocatorValue} to be clickable and click it.");
            if (IsClickable(timeout))
            {
                FindElement(_by).Click();
            }
            else
            {
                throw new TimeoutException($"{LocatorValue} not clickable even after {timeout} seconds");
            }
        }

        /// <summary>
        /// Wait for loading to appear and disappear.
        /// </summary>
        /// <param name="visibilityTimeout">Visibility timeout.</param>
        /// <param name="invisibilityTimeout">Invisibility timeout.</param>
        public void WaitForLoadingToComplete(int visibilityTimeout, int invisibilityTimeout)
        {
            try
            {
                Log.Information($"Waiting for {LocatorValue} to be visible.");
                WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(visibilityTimeout));
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));
            }
            catch
            {
            }
            finally
            {
                Log.Information($"Waiting for {LocatorValue} to disappear.");
                WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(invisibilityTimeout));
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(_by));
            }
        }

        /// <summary>
        /// Wait for custom page refresh.
        /// </summary>
        /// <param name="invisibilityTimeout">Visibility timeout.</param>
        /// <param name="visibilityTimeout">Invisibility timeout.</param>
        public void WaitForRefresh(int invisibilityTimeout, int visibilityTimeout)
        {
            try
            {
                Log.Information($"Waiting for {LocatorValue} to be visible.");
                WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(invisibilityTimeout));
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(_by));
            }
            catch
            {
            }
            finally
            {
                Log.Information($"Waiting for {LocatorValue} to disappear.");
                WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(visibilityTimeout));
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));
            }
        }

        /// <summary>
        /// Wait for UI Element to be visible with custom timeout.
        /// </summary>
        /// <param name="customTimeOut">Custom timeout.</param>
        public void WaitForElementToBeVisible(int customTimeOut)
        {
            try
            {
                Log.Information($"Waiting for {LocatorValue} to be visible.");
                WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(customTimeOut));
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementIsVisible(_by));
            }
            catch (Exception ex)
            {
                string message = "Timeout Exception while trying to locate element with locator: " + LocatorValue.ToString();
                Log.Error(message);
                Log.Error("Exception: " + ex.Message + " : Trace: " + ex.StackTrace);
                throw new TimeoutException(message);
            }
        }

        /// <summary>
        /// Wait for UI Element to be clickable with custom timeout.
        /// </summary>
        /// <param name="customTimeOut">Custom timeout.</param>
        public void WaitForElementToBeClickable(int customTimeOut)
        {
            Log.Information($"Waiting for {LocatorValue} to be clickable.");
            WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(customTimeOut));
            customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(_by));
        }

        /// <summary>
        /// Wait for UI Element to be clickable.
        /// </summary>
        public void WaitForElementToBeClickable()
        {
            Log.Information($"Waiting for {LocatorValue} to be clickable.");
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(_by));
        }

        /// <summary>
        /// Capture ScreenShot of any element.
        /// </summary>
        /// <param name="filePath">File path.</param>
        public void CaptureElementScreenshot(string filePath)
        {
            IWebElement element = FindElement(_by);
            Screenshot screenshot = ((ITakesScreenshot)_webDriver).GetScreenshot();
            Bitmap image = Image.FromStream(new MemoryStream(screenshot.AsByteArray)) as Bitmap;
            image = image.Clone(new Rectangle(element.Location, element.Size), image.PixelFormat);
            image.Save(filePath);
        }

        /// <summary>
        /// Capture Element ScreenShot inside IFrame.
        /// </summary>
        /// <param name="filePath">File path.</param>
        /// <param name="iframeX">IFrame X axis.</param>
        /// <param name="iframeY">IFrame Y axis.</param>
        /// <param name="iframeWidth">IFrame Width.</param>
        /// <param name="iframeHeight">IFrame Height.</param>
        public void CaptureElementScreenshotInsideIFrame(
            string filePath, int iframeX, int iframeY, int iframeWidth, int iframeHeight)
        {
            IWebElement elementInsideIFrame = GetElement();
            Screenshot screenshot = ((ITakesScreenshot)_webDriver).GetScreenshot();
            Bitmap image = Image.FromStream(new MemoryStream(screenshot.AsByteArray)) as Bitmap;
            Bitmap screenshotOfIFrame = image.Clone(new Rectangle(iframeX, iframeY, iframeWidth, iframeHeight), image.PixelFormat);
            Bitmap elementScreenshot = screenshotOfIFrame.Clone(new Rectangle(
                    elementInsideIFrame.Location.X,
                    elementInsideIFrame.Location.Y,
                    elementInsideIFrame.Size.Width,
                    elementInsideIFrame.Size.Height), image.PixelFormat);
            elementScreenshot.Save(filePath);
        }

        /// <summary>
        /// Scroll UI Element into view.
        /// </summary>
        /// <param name="alignToTop">Align To Top, Default value true.</param>
        public void ScrollIntoView(bool alignToTop = true)
        {
            Log.Information($"Scrolling {LocatorValue} into view.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript($"arguments[0].scrollIntoView({alignToTop.ToString().ToLower()});", FindElement(_by));
        }

        /// <summary>
        /// Scroll (top, left, right, bottom).
        /// </summary>
        /// <param name="x">Pixels on x axis. Example if 10px left value will be -10.</param>
        /// <param name="y">Pixels on y axis. Example if 10px down value will be 10.</param>
        public void Scroll(int x, int y)
        {
            Log.Information($"Scrolling x axis {x}, y axis {y}.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript($"window.scrollBy({x}, {y});");
        }

        /// <summary>
        /// Submit the current form.
        /// </summary>
        public void Submit()
        {
            Log.Information($"Submitting {LocatorValue}.");
            FindElement(_by).Submit();
        }

        /// <summary>
        /// Select by value from dropdown.
        /// </summary>
        /// <param name="value">HTML value of dropdown item.</param>
        public void SelectByValue(string value)
        {
            Log.Information($"Selecting by value: {value} in {LocatorValue}.");
            SelectElement selectElement = new SelectElement(FindElement(_by));
            selectElement.SelectByValue(value);
        }

        /// <summary>
        /// Select by dropdown text.
        /// </summary>
        /// <param name="text">Dropdown item text.</param>
        /// <param name="partialMatch">Default value is false. 
        ///  If true a partial match on the Options list will be performed, otherwise exact match.</param>
        public void SelectByText(string text, bool partialMatch = false)
        {
            Log.Information($"Selecting by text: {text} in {LocatorValue}.");
            SelectElement selectElement = new SelectElement(FindElement(_by));
            selectElement.SelectByText(text, partialMatch);
        }

        /// <summary>
        /// Select by dropdown index.
        /// </summary>
        /// <param name="index">Position of the dropdown item.</param>
        public void SelectByIndex(int index)
        {
            Log.Information($"Selecting by index: {index} in {LocatorValue}.");
            SelectElement selectElement = new SelectElement(FindElement(_by));
            selectElement.SelectByIndex(index);
        }

        /// <summary>
        /// Check if an option exists  in select dropdown.
        /// </summary>
        /// <param name="option">Option to check.</param>
        /// <returns>True if option exists else false.</returns>
        public bool DoesOptionExistInSelect(string option)
        {
            SelectElement selectElement = new SelectElement(FindElement(_by));
            var allOptions = selectElement.Options;
            for (int i = 0; i < allOptions.Count; i++)
            {
                if (allOptions[i].Text.Equals(option))
                {
                    return true;
                }
            }

            return false;
        }

        /// <summary>
        /// Click using JavaScript on UI Element.
        /// </summary>
        public void JSClick()
        {
            Log.Information($"Performing JSClick on {LocatorValue}.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].click();", FindElement(_by));
        }

        /// <summary>
        /// Make element visible using JavaScript.
        /// </summary>
        public void JSVisible()
        {
            Log.Information($"Checking visibility of {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].setAttribute('style', 'visibility: visible;');", FindElement(_by));
        }

        /// <summary>
        /// Make file input type element visible using JavaScript.
        /// </summary>
        public void JSVisibleInputField()
        {
            Log.Information($"Checking visibility of {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].style.display='block';", FindElement(_by));
        }

        /// <summary>
        /// Click using JavaScript on UI Element.
        /// </summary>
        /// <param name="attribute">Attribute to be changed.</param>
        /// <param name="value">Value to be changed.</param>
        public void JSChangeAttribute(string attribute, string value)
        {
            Log.Information($"Changing attribute of {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript($"arguments[0].setAttribute('{attribute}', '" + value + "')", FindElement(_by));
        }

        /// <summary>
        /// Remove the element using JavaScript.
        /// </summary>
        public void JSRemoveElement()
        {
            Log.Information($"Removing element of {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].remove();", FindElement(_by));
        }

        /// <summary>
        /// Remove the attribute using JavaScript.
        /// </summary>
        /// <param name="attribute">Attribute to be removed.</param>
        public void JSRemoveAttribute(string attribute)
        {
            Log.Information($"Removing attribute of {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript($"arguments[0].removeAttribute('{HttpUtility.UrlEncode(attribute)}')", FindElement(_by));
        }

        /// <summary>
        /// GetText using JS.
        /// </summary>
        /// <returns>Text using JS.</returns>
        public string JSGetText()
        {
            Log.Information($"Getting text of element {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            return (string)executor.ExecuteScript("return arguments[0].innerHTML;", FindElement(_by));
        }

        /// <summary>
        /// Check the element.
        /// </summary>
        public void Check()
        {
            Log.Information($"Checking checkbox of {LocatorValue}.");
            var element = FindElement(_by);

            if (!element.Selected)
            {
                element.Click();
            }
        }

        /// <summary>
        /// Uncheck the element.
        /// </summary>
        public void Uncheck()
        {
            Log.Information($"Unchecking checkbox of {LocatorValue}.");
            var element = FindElement(_by);

            if (element.Selected)
            {
                element.Click();
            }
        }

        /// <summary>
        /// Replace ETL Query.
        /// </summary>
        /// <param name="query">ETL Job Query.</param>
        public void SetCodeMirrorQuery(string query)
        {
            Log.Information($"Setting query in {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript("arguments[0].CodeMirror.setValue(`" + query + "`);", FindElement(_by));
        }

        /// <summary>
        /// Get the particular attribute value.
        /// </summary>
        /// <param name="attribute">HTML attribute.</param>
        /// <returns>String value of attribute.</returns>
        public string GetAttributeValue(string attribute)
        {
            Log.Information($"Getting {attribute} of {LocatorValue}.");
            return FindElement(_by).GetAttribute(attribute);
        }

        /// <summary>
        /// Wait for UI Element to disappear.
        /// </summary>
        public void WaitForElementToDisappear()
        {
            Log.Information($"Waiting for {LocatorValue} to disappear.");
            _wait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(_by));
        }

        /// <summary>
        /// Wait for UI Element to disappear with custom timeout.
        /// </summary>
        /// <param name="customTimeOut">Custom timeout.</param>
        public void WaitForElementToDisappear(int customTimeOut)
        {
            try
            {
                Log.Information($"Waiting for {LocatorValue} to disappear.");
                WebDriverWait customWait = new WebDriverWait(_webDriver, TimeSpan.FromSeconds(customTimeOut));
                customWait.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.InvisibilityOfElementLocated(_by));
            }
            catch (TimeoutException ex)
            {
                Log.Error("Timeout Exception while checking disappearance of element with locator: " + LocatorValue.ToString());
                Log.Error("Exception: " + ex.Message + " : Trace: " + ex.StackTrace);
                throw;
            }
        }

        /// <summary>
        /// Get all element as a List if more than one element exists.
        /// </summary>
        /// <returns>List of UI Elements.</returns>
        public List<UIElement> GetElements()
        {
            Log.Information($"Getting list of elements with locator {LocatorValue}");
            List<UIElement> uiElements = new List<UIElement>();

            int i = 1;
            foreach (var element in FindElements(_by))
            {
                if (!LocatorValue.StartsWith("("))
                {
                    LocatorValue = "(" + LocatorValue;
                }

                if (!LocatorValue.EndsWith(")"))
                {
                    LocatorValue = LocatorValue + ")";
                }

                uiElements.Add(new UIElement(_webDriver, "xpath", string.Format(LocatorValue + "[{0}]", i), _timeout));
                i++;
            }

            return uiElements;
        }

        /// <summary>
        /// Get UI Element.
        /// </summary>
        /// <returns>UI Element.</returns>
        public IWebElement GetElement()
        {
            Log.Information($"Get element with locator: {LocatorValue}.");
            IWebElement webElement = FindElement(_by);
            return webElement;
        }

        /// <summary>
        /// Get list of WebElements.
        /// </summary>
        /// <param name="by">By object of locator.</param>
        /// <returns>List of WebElement.</returns>
        public IReadOnlyCollection<IWebElement> FindElements(By by)
        {
            return _wait.Until(webDriver => webDriver.FindElements(by));
        }

        /// <summary>
        /// Find element using the by object.
        /// </summary>
        /// <param name="by">By object of locator.</param>
        /// <returns>IWebElement.</returns>
        public IWebElement FindElement(By by)
        {
            return _wait.Until(webDriver => webDriver.FindElement(by));
        }

        /// <summary>
        /// Set innerText using JS.
        /// </summary>
        /// <param name="value">Value to set in element.</param>
        public void SetInnerText(string value)
        {
            Log.Information($"Setting innerText of element {LocatorValue} using JS.");
            IJavaScriptExecutor executor = (IJavaScriptExecutor)_webDriver;
            executor.ExecuteScript($"arguments[0].innerText = '{value}';", FindElement(_by));
        }

        /// <summary>
        /// Get the By object of locator.
        /// </summary>
        /// <returns>By object.</returns>
        private By GetBy()
        {
            By by;

            switch (LocatorType)
            {
                case FindByLocatorConstants.Id:
                    by = By.Id(LocatorValue);
                    break;
                case FindByLocatorConstants.Xpath:
                    by = By.XPath(LocatorValue);
                    break;
                case FindByLocatorConstants.Css:
                    by = By.CssSelector(LocatorValue);
                    break;
                case FindByLocatorConstants.ClassName:
                    by = By.ClassName(LocatorValue);
                    break;
                case FindByLocatorConstants.LinkText:
                    by = By.LinkText(LocatorValue);
                    break;
                case FindByLocatorConstants.TagName:
                    by = By.TagName(LocatorValue);
                    break;
                case FindByLocatorConstants.Name:
                    by = By.Name(LocatorValue);
                    break;
                case FindByLocatorConstants.PartialLinkText:
                    by = By.PartialLinkText(LocatorValue);
                    break;
                default:
                    throw new Exception($"Types allowed are: id, xpath, css, class, linktext, tagname, " +
                        $"name, partiallinktext. Type found is {LocatorType}");
            }

            return by;
        }
    }
}
