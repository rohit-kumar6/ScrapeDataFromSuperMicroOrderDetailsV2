namespace Automation.Core.Web
{
    using System.Collections.Generic;

    /// <summary>
    /// Plain old object to convert entire locator JSON file to an object.
    /// </summary>
    internal class LocatorJSONObject
    {
        /// <summary>
        /// Gets or sets Page Object name.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets list of Page Element objects from Page Object locators JSON file.
        /// </summary>
        public List<PageElement> Elements { get; set; }
    }
}
