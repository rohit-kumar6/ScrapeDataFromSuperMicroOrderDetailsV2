using Automation.Core;
using System;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Windows;
using System.Windows.Controls;
using System.Collections.Generic;
using OpenQA.Selenium;
using Serilog;

namespace Automation.SuperNova
{
    /// <summary>
    /// Interaction logic for UserControlSuperNova.xaml
    /// </summary>
    public partial class UserControlSuperNova : UserControl
    {
        /// <summary>
        ///  UserControlHire Main.
        /// </summary>
        internal static UserControlSuperNova _userControlHire;
        private readonly InputObject _ipObj;
        private BackgroundWorker _backgroundWorker;

        public UserControlSuperNova()
        {
            InitializeComponent();
            _userControlHire = this;
            _ipObj = new InputObject();
        }

        private void SuperNova_Start(object sender, RoutedEventArgs e)
        {
            startSuperNova.IsEnabled = false;
            IsEnabled = false;
            App.GetWindowInstance().ToggleToolsMenuView();

            _ipObj.startDate = Convert.ToDateTime(startDate.SelectedDate.ToString()).ToString("MM/dd/yyyy");
            _ipObj.endDate = Convert.ToDateTime(endDate.SelectedDate.ToString()).ToString("MM/dd/yyyy");

            var coustomerIdList = new List<string>();
            foreach (var item in coustomerId.SelectedItems)
            {
                var itemString = item.ToString();
                var startIndex = itemString.LastIndexOf(":") + 2;
                var stringSize = itemString.Length - startIndex;
                coustomerIdList.Add(itemString.Substring(startIndex, stringSize));
            }

            var orderTypeList = new List<string>();
            foreach (var item in orderType.SelectedItems)
            {
                var itemString = item.ToString();
                var startIndex = itemString.LastIndexOf(":") + 2;
                var stringSize = itemString.Length - startIndex;
                orderTypeList.Add(itemString.Substring(startIndex, stringSize));
            }

            _ipObj.customerIdList = coustomerIdList;
            _ipObj.orderTypeList = orderTypeList;
            _ipObj.outputPath = outputPath.Text;

            _backgroundWorker = new BackgroundWorker();
            _backgroundWorker.DoWork += Worker_DoWork;
            _backgroundWorker.RunWorkerAsync(_ipObj);
            _backgroundWorker.RunWorkerCompleted += Worker_RunWorkerCompleted;
        }

        private void DatePicker_SelectedDateChanged(object sender, RoutedEventArgs e)
        {
            var picker = sender as DatePicker;
            _ = picker.SelectedDate;
        }

        private void Worker_RunWorkerCompleted(object sender, RunWorkerCompletedEventArgs e)
        {
            startSuperNova.IsEnabled = true;
            IsEnabled = true;
            App.GetWindowInstance().ToggleToolsMenuView();

            if (e.Error != null)
            {
                MessageBox.Show(
                    "Tool Execution Fail",
                    "Run Status",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error,
                    MessageBoxResult.OK);
            }
            else
            {
                MessageBox.Show(
                    "Tool Execution Pass",
                    "Run Status",
                    MessageBoxButton.OK,
                    MessageBoxImage.Information,
                    MessageBoxResult.OK);
            }

            _backgroundWorker.Dispose();
        }

        private void BrowsebtnOpenOutputFile_Click(object sender, RoutedEventArgs e)
        {
            System.Windows.Forms.FolderBrowserDialog openFolderDialog = new System.Windows.Forms.FolderBrowserDialog
            {
                ShowNewFolderButton = true,
            };
            _ = openFolderDialog.ShowDialog();
            outputPath.Text = openFolderDialog.SelectedPath;
            openFolderDialog.Dispose();
        }

        private void OpenbtnOutput_Click(object sender, RoutedEventArgs e)
        {
            if (Directory.Exists(outputPath.Text))
            {
                Process.Start(outputPath.Text);
            }
        }

        private void Worker_DoWork(object sender, DoWorkEventArgs e)
        {
            var ipObj = e.Argument as InputObject;
            string filePath = $@"{ipObj.outputPath}\AutomationLog.log";
            Log.Logger = new LoggerConfiguration()
                    .MinimumLevel.Debug()
                      .WriteTo.File(
                        filePath,
                        rollingInterval: RollingInterval.Day,
                        rollOnFileSizeLimit: true,
                        fileSizeLimitBytes: 123456)
                      .CreateLogger();

            IWebDriver webDriver = null;
            try
            {
                Log.Information("Process started");
                webDriver = WebDriverUtils.DriverSetup(
                    headless: false,
                    customTimeout: 2);
                new ProcessActivity(webDriver, ipObj).Execute();
            }
            catch (Exception ex)
            {
                Log.Error(ex.Message + " " + ex.StackTrace);
                throw;
            }
            finally
            {
                Log.Information("Web driver dispose");
                webDriver?.Dispose();
            }
        }
    }
}