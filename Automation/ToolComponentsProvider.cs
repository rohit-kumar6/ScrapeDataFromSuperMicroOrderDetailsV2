namespace Automation
{
    using System.Collections.Generic;
    using Automation.SuperNova;
    using MaterialDesignThemes.Wpf;

    /// <summary>
    /// Utility class to create tool components for tools.
    /// </summary>
    public static class ToolComponentsProvider
    {
        /// <summary>
        /// Creates the tools list present in Taskless Bot.
        /// </summary>
        /// <returns>List of Tool Components.</returns>
        public static List<ToolComponent> GetToolComponentList()
        {
            List<ToolComponent> toolComponentList = new List<ToolComponent>
            {
                GetSuperNova(),
            };

            return toolComponentList;
        }

        private static ToolComponent GetSuperNova()
        {
            return new ToolComponent
            {
                ToolUserControlType = typeof(UserControlSuperNova),
                ToolName = "Super Nova",
                ToolTag = "Super Nova",
                ToolIcon = PackIconKind.Latest,
            };
        }
    }
}
