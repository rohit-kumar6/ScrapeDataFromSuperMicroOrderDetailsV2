using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using Serilog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using WebDriverManager;
using WebDriverManager.DriverConfigs.Impl;
using WebDriverManager.Helpers;

namespace Automation.Core
{
    /// <summary>
    /// Utility class for WebDriver.
    /// </summary>
    public static class WebDriverUtils
    {
        /// <summary>
        /// Returns a Web Driver.
        /// </summary>
        /// <param name="chromeDriver">Chrome driver.</param>
        /// <param name="headless">True is headless mode preferred.</param>
        /// <param name="optionalDownloadPath">File download path from browser.</param>
        /// <param name="customTimeout">Custom timeout of driver.</param>
        /// <param name="preserveCookies">Preserve Cookies.</param>
        /// <param name="resetZoomPercentageUrlSet">Reset Zoom Percentage URL Set.</param>
        /// <param name="sourceProfileName">Source Profile name.</param>
        /// <param name="unhandledPromptBehavior">Behavior to handle alerts.</param>
        /// <returns>WebDriver Instance of chrome.</returns>
        public static IWebDriver DriverSetup(
            bool headless = false,
            string optionalDownloadPath = "",
            int customTimeout = 0,
            UnhandledPromptBehavior unhandledPromptBehavior = UnhandledPromptBehavior.Default)
        {
            Log.Information("Launching the Chrome Instance.");
            ChromeOptions options = new ChromeOptions();
            options.AddUserProfilePreference("plugins.always_open_pdf_externally", true);
            options.AddArgument("--log-level=3");
            options.AddArguments("--start-maximized");
            if (!unhandledPromptBehavior.Equals(UnhandledPromptBehavior.Default))
            {
                options.UnhandledPromptBehavior = unhandledPromptBehavior;
            }

            if (headless)
            {
                options.AddArgument("headless");
            }

            if (!string.IsNullOrEmpty(optionalDownloadPath))
            {
                options.AddUserProfilePreference("download.default_directory", optionalDownloadPath);
                options.AddUserProfilePreference("download.prompt_for_download", "false");
                options.AddUserProfilePreference("profile.default_content_setting_values.automatic_downloads", 1);
            }

            new DriverManager().SetUpDriver(new ChromeConfig(), VersionResolveStrategy.MatchingBrowser);
            ChromeDriverService chromeDriverService = ChromeDriverService.CreateDefaultService();
            chromeDriverService.HideCommandPromptWindow = true;
            ChromeDriver webDriver = customTimeout > 0 ?
                new ChromeDriver(chromeDriverService, options, TimeSpan.FromMinutes(customTimeout)) :
                new ChromeDriver(chromeDriverService, options);
            return webDriver;
        }
    }
}
