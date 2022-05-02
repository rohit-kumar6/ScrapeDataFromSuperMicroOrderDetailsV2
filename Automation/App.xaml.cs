namespace Automation
{
    using System.Windows;

    /// <summary>
    /// Interaction logic for App.xaml file.
    /// </summary>
    public partial class App : Application
    {
        private static MainWindow window;

        /// <summary>
        /// Gets the Window Instance object.
        /// </summary>
        /// <returns>Window object.</returns>
        public static MainWindow GetWindowInstance()
        {
            return window;
        }

        /// <summary>
        /// On Start up method.
        /// </summary>
        /// <param name="e">Event argument.</param>
        protected override void OnStartup(StartupEventArgs e)
        {
            MainWindowComponent mainWindowComponent = new MainWindowComponent
            {
                ApplicationName = "Automation Bot",
                WindowTitle = "Automation Bot",
                HomeUserControl = new UserControlHome(),
                ToolList = ToolComponentsProvider.GetToolComponentList(),
            };

            window = new MainWindow(mainWindowComponent);
            window.Show();
        }
    }
}