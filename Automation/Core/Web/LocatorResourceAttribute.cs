namespace Automation.Core.Web
{
    using System;

    /// <summary>
    /// Annotation for Locators file for UI Automation Framework.
    /// </summary>
    [AttributeUsage(AttributeTargets.Class, AllowMultiple = false)]
    public class LocatorResourceAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LocatorResourceAttribute"/> class.
        /// </summary>
        /// <param name="resourceType">Actual resx filename.</param>
        /// <param name="resourceName">Resource name contained inside resx.</param>
        public LocatorResourceAttribute(Type resourceType, string resourceName)
        {
            this.Content = ResourceUtils.GetResourceFileContent(resourceType, resourceName);
        }

        /// <summary>
        /// Gets or sets resource content.
        /// </summary>
        public string Content { get; set; }
    }
}
