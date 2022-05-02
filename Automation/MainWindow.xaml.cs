using MaterialDesignThemes.Wpf;
using Serilog;
using System;
using System.Collections.Generic;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Input;

namespace Automation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {

        private const string HomeListViewTag = "Home";
        private const string SearchListViewTag = "Search";
        private readonly UserControl homeUserControl;
        private readonly Dictionary<string, KeyValuePair<Type, object[]>> toolMapping = new Dictionary<string, KeyValuePair<Type, object[]>>();

        /// <summary>
        /// Initializes a new instance of the <see cref="MainWindow"/> class.
        /// </summary>
        /// <param name="mainWindowComponent">Main Window Component.</param>
        public MainWindow(MainWindowComponent mainWindowComponent)
        {
            this.homeUserControl = mainWindowComponent.HomeUserControl;
            this.InitializeComponent();
            this.Window.Title = mainWindowComponent.WindowTitle;
            this.ApplicationName.Text = mainWindowComponent.ApplicationName;
            this.ApplicationName.Text = mainWindowComponent.ApplicationName;
            AddTools(this.ListViewMenu, mainWindowComponent.ToolList, toolMapping);
        }

        public static void AddTools(
            ListView listViewMenu,
            List<ToolComponent> toolComponentList,
            Dictionary<string, KeyValuePair<Type, object[]>> toolMapping)
        {
            foreach (ToolComponent tool in toolComponentList)
            {
                listViewMenu.Items.Add(GetListViewItemOfTool(tool));
                toolMapping.Add(tool.ToolTag, new KeyValuePair<Type, object[]>(tool.ToolUserControlType, tool.ToolUserControlParams));
            }
        }

        private static ListViewItem GetListViewItemOfTool(ToolComponent tool)
        {
            PackIcon icon = new PackIcon
            {
                Margin = new Thickness(10),
                Kind = tool.ToolIcon,
                VerticalAlignment = VerticalAlignment.Center,
                Width = 25,
                Height = 25,
            };

            TextBlock textBlock = new TextBlock
            {
                Text = tool.ToolName,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(20, 10, 20, 10),
            };

            StackPanel stackPanel = new StackPanel
            {
                Orientation = Orientation.Horizontal,
            };
            stackPanel.Children.Add(icon);
            stackPanel.Children.Add(textBlock);

            return new ListViewItem
            {
                Height = 60,
                Tag = tool.ToolTag,
                Content = stackPanel,
            };
        }


        /// <summary>
        /// Toggle tools menu view.
        /// </summary>
        public void ToggleToolsMenuView()
        {
            this.ToolsPanel.IsEnabled = this.ToolsPanel.IsEnabled ? false : true;
        }

        private async void Close_Click(object sender, RoutedEventArgs e)
        {
            Log.Information("Application Shutdown Triggered");
            Application.Current.Shutdown();
            Log.Information("Application Shutdown");
        }

        private void Minimize_Click(object sender, RoutedEventArgs e)
        {
            this.WindowState = WindowState.Minimized;
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (e.RightButton == MouseButtonState.Pressed)
            {
                return;
            }

            this.DragMove();
        }

        private void ButtonOpenMenu_Click(object sender, RoutedEventArgs e)
        {
            this.ButtonCloseMenu.Visibility = Visibility.Visible;
            this.ButtonOpenMenu.Visibility = Visibility.Collapsed;
            this.ExpandedMenuIconPanel.Visibility = Visibility.Visible;
            this.OpenMenuSearch.Visibility = Visibility.Collapsed;
            this.Search.Focus();
        }

        private void ButtonCloseMenu_Click(object sender, RoutedEventArgs e)
        {
            this.ButtonCloseMenu.Visibility = Visibility.Collapsed;
            this.ButtonOpenMenu.Visibility = Visibility.Visible;
            this.ExpandedMenuIconPanel.Visibility = Visibility.Collapsed;
            this.OpenMenuSearch.Visibility = Visibility.Visible;
        }

        private void Search_Click(object sender, RoutedEventArgs e)
        {
            this.ButtonOpenMenu.RaiseEvent(new RoutedEventArgs(Button.ClickEvent));
        }

        private void Search_TextChanged(object sender, RoutedEventArgs e)
        {
            TextBox searchBox = sender as TextBox;
            string searchedText = searchBox.Text.ToLower().Trim();
            HandleSearchTextChangeEvent(this.ListViewMenu, HomeListViewTag, SearchListViewTag, searchedText);
        }

        public static void HandleSearchTextChangeEvent(
            ListView toolsListView, string homeListViewTag, string searchListViewTag, string searchedText)
        {
            foreach (ListViewItem tool in toolsListView.Items)
            {
                string toolName = tool.Tag.ToString();
                bool showTool = toolName.Equals(homeListViewTag)
                             || toolName.Equals(searchListViewTag)
                             || toolName.ToLower().Contains(searchedText);
                tool.Visibility = showTool ? Visibility.Visible : Visibility.Collapsed;
            }
        }

        private void ListViewMenu_SelectionChanged(object sender, SelectionChangedEventArgs e)
        {
            ListViewItem selectedItem = (ListViewItem)this.ListViewMenu.SelectedItem;
            string selectedToolTag = selectedItem.Tag.ToString();
            if (selectedToolTag.Equals(HomeListViewTag))
            {
                this.GridPrincipal.Children.Clear();
                this.GridPrincipal.Children.Add(homeUserControl);
            }
            else if (this.toolMapping.ContainsKey(selectedToolTag))
            {
                this.GridPrincipal.Children.Clear();
                this.GridPrincipal.Children.Add((UserControl)Activator.CreateInstance(toolMapping[selectedToolTag].Key, toolMapping[selectedToolTag].Value));
            }
        }
    }
}
