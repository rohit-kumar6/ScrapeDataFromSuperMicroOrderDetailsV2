namespace Automation.Core.Web
{
    /// <summary>
    /// Plain Old object for creating an object out of json block.
    /// </summary>
    internal class PageElement
    {
        /// <summary>
        /// Gets or sets PageElement name from locator json.
        /// </summary>
        public string Name { get; set; }

        /// <summary>
        /// Gets or sets PageElement type from locator json.
        /// </summary>
        public string Type { get; set; }

        /// <summary>
        /// Gets or sets PageElement Value from locator json.
        /// </summary>
        public string Value { get; set; }
    }
}
