﻿
using NetOffice.OutlookApi;
using NetOffice.OutlookApi.Enums;
using System;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Input;

namespace readClashReport.Information.UI
{
    /// <summary>
    /// Interaction logic for InformationUI.xaml
    /// </summary>
    public partial class InformationUI : Window
    {
        public InformationUI()
        {
            InitializeComponent();
            this.versionLabel.Content = VersionInfo.version;
            this.aboutReleaseInfo.Text = String.Join(Environment.NewLine, VersionInfo.aboutReleaseInfo);
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {

            if (Mouse.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void CloseBtn_Click(object sender, RoutedEventArgs e)
        {
            this.Close();
        }

 
        private async void reportBtn_Click(object sender, RoutedEventArgs e)
        {
            //this.Close();
            //NetOffice.OutlookApi.Application outApp = new NetOffice.OutlookApi.Application();
            //MailItem mailItem = (MailItem)outApp.CreateItem(OlItemType.olMailItem);
            //mailItem.Subject = "Clash Report Viewer Issue";
            //mailItem.Body = "**Describe Issue Here**";
            //mailItem.To = "dodzi@windowslive.com";
            //mailItem.Importance = OlImportance.olImportanceNormal;

            //await Task.Run(() =>
            //{

            //    try
            //    {
            //        mailItem.Display();

            //    }
            //    catch (System.Exception ex)
            //    {

            //        Debug.WriteLine(ex.Message);
            //    }


            //});

        }

        private void connectBtn_Click(object sender, RoutedEventArgs e)
        {
            // System.Diagnostics.Process.Start("https://www.linkedin.com/in/dodziagbenorku/");
            OpenBrowser("https://www.linkedin.com/in/dodziagbenorku/");
        }

        public static void OpenBrowser(string url)
        {

            try
            {
                Process.Start(url);
            }
            catch
            {
                // hack because of this: https://github.com/dotnet/corefx/issues/10361
                if (RuntimeInformation.IsOSPlatform(OSPlatform.Windows))
                {
                    url = url.Replace("&", "^&");
                    Process.Start(new ProcessStartInfo("cmd", $"/c start {url}") { CreateNoWindow = true });
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.Linux))
                {
                    Process.Start("xdg-open", url);
                }
                else if (RuntimeInformation.IsOSPlatform(OSPlatform.OSX))
                {
                    Process.Start("open", url);
                }
                else
                {
                    throw;
                }
            }
        }
    }
}
