namespace Automation.Core.Web
{
    using System;

    /// <summary>
    /// Annotation for Locators for UI Automation Framework.
    /// </summary>
    [AttributeUsage(AttributeTargets.Field, AllowMultiple = true)]
    public class LocatorAttribute : Attribute
    {
        /// <summary>
        /// Initializes a new instance of the <see cref="LocatorAttribute"/> class.
        /// </summary>
        /// <param name="name">Unique locator name.</param>
        public LocatorAttribute(string name)
        {
            this.Name = name;
        }

        /// <summary>
        /// Gets locator unique name.
        /// </summary>
        public string Name { get; }
    }
}
