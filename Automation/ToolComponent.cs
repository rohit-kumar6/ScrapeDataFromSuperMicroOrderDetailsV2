namespace Automation
{
    using System;
    using System.ComponentModel.DataAnnotations;
    using MaterialDesignThemes.Wpf;

    /// <summary>
    /// POCO for Tool Component.
    /// </summary>
    public class ToolComponent
    {
        /// <summary>
        /// Gets or sets the Tool Name.
        /// </summary>
        [Required]
        public string ToolName { get; set; }

        /// <summary>
        /// Gets or sets the Tool Tag.
        /// </summary>
        [Required]
        public string ToolTag { get; set; }

        /// <summary>
        /// Gets or sets the Tool Icon.
        /// </summary>
        [Required]
        public PackIconKind ToolIcon { get; set; }

        /// <summary>
        /// Gets or sets the typeof Tool UserControl.
        /// </summary>
        [Required]
        public Type ToolUserControlType { get; set; }

        /// <summary>
        /// Gets or sets the Tool UserControl Constructor parameters.
        /// </summary>
        public object[] ToolUserControlParams { get; set; }
    }
}
