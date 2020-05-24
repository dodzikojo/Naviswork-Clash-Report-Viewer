﻿using Ookii.Dialogs.Wpf;
using PuppeteerSharp;
using readClashReport.Properties;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Data;
using System.Windows.Input;
using System.Windows.Interop;

namespace readClashReport
{
    public class htmlFiles
    {
        public string filename { get; set; }
        public string clashes { get; set; }
        public string newClashes { get; set; }
        public string active { get; set; }
        public string reviewed { get; set; }
        public string type { get; set; }
        public static string[,] data { get; set; }

    }

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int GWL_STYLE = -16;
        private const int WS_SYSMENU = 0x80000;

        public static string location { get; set; }
        private static string filename { get; set; }
        public static string currentItem { get; set; }
        public static List<htmlFiles> fileData = new List<htmlFiles>();
        public static List<string> filenamesList = new List<string>();
        public static List<string> clashesList = new List<string>();
        public static List<string> activeList = new List<string>();
        public static List<string> reviewedList = new List<string>();
        public static List<string> approvedList = new List<string>();
        public static List<string> typeList = new List<string>();
        public static List<string> resolvedList = new List<string>();
        public static List<string> toleranceList = new List<string>();
        public static List<string> newList = new List<string>();
        List<string> filePathList = new List<string>();
        public static bool openExcelBool { get; set; }



        /// <summary>
        /// Window Object
        /// </summary>
        public MainWindow()
        {
            //InitializeComponent();

            this.Dispatcher.Invoke(() =>
            {
                InitializeComponent();

                Debug.WriteLine(Settings.Default.WindowLocation);
                //Set window location
                if (Settings.Default.WindowLocation != null)
                {
                    this.Top = Settings.Default.WindowLocation.Y;
                    this.Left = Settings.Default.WindowLocation.X;
                }

                //Set window size
                if (Settings.Default.WindowSize != null)
                {
                    this.Width = Settings.Default.WindowSize.Width;
                    this.Height = Settings.Default.WindowSize.Height;
                }

            });

            try

            {
                if (Properties.Settings.Default.filePathSetting.Length > 0)
                {
                    //this.folderTxtBox.Text = Properties.Appsettings.Default.filePathSetting;
                    getHTMLfiles(this.folderTxtBox.Text = Properties.Settings.Default.filePathSetting);
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }

        }

        /// <summary>
        /// Gets the html files in the user specified folder.
        /// </summary>
        /// <param name="path"></param>
        /// <returns></returns>
        private List<string> getHTMLfiles(string path)
        {
            List<string> filePathListTemp = new List<string>();
            string[] fileArray = Directory.GetFiles(path, "*.html");
            location = path;
            var results = isValid(fileArray);

            if (results.Item2)
            {
                this.folderTxtBox.Text = path;
                //htmlFiles.data = HTMLdata2DArr(results.Item1);
                DisplayData(results.Item1);

            }
            return filePathListTemp;
        }

        /// <summary>
        /// Takes HTML data and creates a 2D Array
        /// </summary>
        /// <param name="data"></param>
        /// <returns></returns>
        private string[,] HTMLdata2DArr(List<string> data)
        {
            int num = 0;
            string[,] tempData = new string[data.Count, 9];
            Debug.WriteLine($"data count is {data.Count}");
            try
            {
                foreach (string filename in data)
                {

                    try
                    {
                        tempData[num, 0] = filename;
                        tempData[num, 1] = toleranceList[num];
                        tempData[num, 2] = clashesList[num];
                        tempData[num, 3] = newList[num];
                        tempData[num, 4] = activeList[num];
                        tempData[num, 5] = reviewedList[num];
                        tempData[num, 6] = approvedList[num];
                        tempData[num, 7] = resolvedList[num];
                        tempData[num, 8] = typeList[num];

                    }
                    catch (Exception ex1)
                    {

                        Debug.WriteLine(ex1.Message);
                    }

                    num++;
                }
                MainWindow.filenamesList.Clear();

            }
            catch (Exception e)
            {

                Debug.WriteLine(e.Message);
            }


            return tempData;
        }

        /// <summary>
        /// Browse button click event
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void browseBtn_Click(object sender, RoutedEventArgs e)
        {
            VistaFolderBrowserDialog folderBrowser = new VistaFolderBrowserDialog();
            folderBrowser.Description = "Choose folder where Navisworks clash reports are stored";
            folderBrowser.UseDescriptionForTitle = true;
            Nullable<bool> fdRun = folderBrowser.ShowDialog();
            try
            {
                if (fdRun == true)
                {
                    string chosenPath = folderBrowser.SelectedPath.ToString();
                    getHTMLfiles(chosenPath);

                }
            }
            catch (Exception ex1)
            {
                Debug.WriteLine(ex1.Message);
            }
        }

        /// <summary>
        /// Checks if the file is a valid html and clash report file.
        /// </summary>
        /// <param name="tempFilepaths"></param>
        /// <returns></returns>
        public (List<string>, bool) isValid(string[] tempFilepaths)
        {
            bool valid;
            List<string> filePathListTemp = new List<string>();
            filePathList.Clear();
            if (tempFilepaths.Length > 0)
            {
                if (tempFilepaths[0].ToString().ToLower().Contains("vs"))
                {
                    fileData.Clear();
                    valid = true;
                    this.countLabel.Content = tempFilepaths.Length.ToString();
                    foreach (string item in tempFilepaths)
                    {
                        if (Path.GetFileNameWithoutExtension(item).ToLower().Contains("vs") && Path.GetExtension(item).ToString() == ".html")
                        {
                            filePathListTemp.Add(item);
                        }
                    }
                }
                else
                {
                    valid = false;
                    MessageBox.Show("No valid files found in chosen folder." +
                        " Check that folder contains Navisworks clash test reports." +
                        " For support, contact the developer.", "Error",
                        MessageBoxButton.OK,
                        MessageBoxImage.Error,
                        MessageBoxResult.OK);
                }
            }
            else
            {
                valid = false;
                MessageBox.Show("No valid files found in chosen folder." +
                    " Check that folder contains Navisworks clash test reports." +
                    " For support, contact the developer.", "Error",
                    MessageBoxButton.OK,
                    MessageBoxImage.Error,
                    MessageBoxResult.OK);
            }

            return (filePathListTemp, valid);
        }

        /// <summary>
        /// Displays .html files in the listview panel.
        /// </summary>
        public async void DisplayData(List<string> filepathlist)
        {
            filesListView.ItemsSource = null;
            filesListView.Items.Clear();
            fileData.Clear();
            //htmlFiles.data = null;
            await Task.Run(() =>
            {
                try
                {
                    foreach (string html in filepathlist)
                    {
                        readHTML.readHTMLData(@html);
                        this.Dispatcher.Invoke(() =>
                        {
                            filesListView.ItemsSource = fileData;
                            filesListView.Items.Refresh();

                        });
                    }
                    this.Dispatcher.Invoke(() =>
                    {
                        this.webViewer.Navigate(filepathlist[0]);
                        currentItem = filepathlist[0];

                    });
                    htmlFiles.data = HTMLdata2DArr(filenamesList);

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
            });

            try
            {
                CollectionView view = (CollectionView)CollectionViewSource.GetDefaultView(filesListView.ItemsSource);
                view.Filter = UserFilter;
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

        static async void getAllLinks(string file)
        {
            LaunchOptions options = new LaunchOptions();
            using (var browser = await Puppeteer.LaunchAsync(options))
            using (var page = await browser.NewPageAsync())
            {
                await page.GoToAsync(file);
                var jsSelectAllAnchors = @"Array.from(document.querySelectorAll('a')).map(a => a.href);";
                var urls = await page.EvaluateExpressionAsync<string[]>(jsSelectAllAnchors);
                foreach (string url in urls)
                {
                    Debug.WriteLine($"Url: {url}");
                }
                //Console.WriteLine("Press any key to continue...");
                //Console.ReadLine();
            }
        }

        [DllImport("user32.dll", SetLastError = true)]
        private static extern int GetWindowLong(IntPtr hWnd, int nIndex);
        [DllImport("user32.dll")]
        private static extern int SetWindowLong(IntPtr hWnd, int nIndex, int dwNewLong);

        private bool UserFilter(object item)
        {
            if (String.IsNullOrEmpty(txtFilter.Text))
                return true;
            else
                return ((item as htmlFiles).filename.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void txtFilter_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            CollectionViewSource.GetDefaultView(filesListView.ItemsSource).Refresh();
        }


        private void filesListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {
            foreach (htmlFiles item in this.filesListView.SelectedItems)
            {
                //Debug.WriteLine(Path.Combine(this.folderTxtBox.Text, item.filename + ".html"));
                currentItem = Path.Combine(this.folderTxtBox.Text, item.filename + ".html");
                this.webViewer.Navigate(currentItem);
            }
        }



        /// <summary>
        /// Window closing event.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {


            // Copy window location to app settings
            Settings.Default.WindowLocation = new System.Drawing.Point(Convert.ToInt32(this.Left), Convert.ToInt32(this.Top));

            Debug.WriteLine(this.folderTxtBox.Text);

            //Copy window size to app settings
            if (this.WindowState == WindowState.Normal)
            {
                Settings.Default.WindowSize = new System.Drawing.Size(Convert.ToInt32(this.Width), Convert.ToInt32(this.Height));
            }
            else
            {
                Settings.Default.WindowSize = new System.Drawing.Size(Convert.ToInt32(this.RestoreBounds.Size.Width), Convert.ToInt32(this.RestoreBounds.Size.Height));
            }
            Properties.Settings.Default.Save();
            e.Cancel = true;
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                Environment.Exit(0);
                Debug.WriteLine(ex.Message);
            }
        }


        /// <summary>
        /// TODO: Send .html as a PDF via email.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        //private void emailBtn_Click(object sender, RoutedEventArgs e)
        //{
        //    Debug.WriteLine("this is the email button.");

        //}

        static async void createPDF(string file)
        {
            await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision);
            var browser = await Puppeteer.LaunchAsync(new LaunchOptions
            {
                Headless = true
            });
            var page = await browser.NewPageAsync();
            await page.SetViewportAsync(new ViewPortOptions { Width = 50000, Height = 50 });
            await page.GoToAsync(file);

            PdfOptions options = new PdfOptions();
            options.Width = 2000;
            options.Height = 1200;
            options.MarginOptions.Left = "100";
            options.MarginOptions.Right = "100";

            await page.PdfAsync(Path.Combine(location, Path.GetFileNameWithoutExtension(file) + ".pdf"), options);



        }

        /// <summary>
        /// TODO: Save the .html file as PDF.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void excelBtn_Click(object sender, RoutedEventArgs e)
        {
            //Debug.WriteLine("this is the save pdf button");

            excel.writeExcel.writeExcelFile(htmlFiles.data);
        }

        private void openExcelBtn_Checked(object sender, RoutedEventArgs e)
        {
            openExcelBool = true;
        }

        private void openExcelBtn_Unchecked(object sender, RoutedEventArgs e)
        {
            openExcelBool = false;
        }

        private void aboutBtn_Click(object sender, RoutedEventArgs e)
        {
            Information.UI.InformationUI aboutWindow = new Information.UI.InformationUI();
            aboutWindow.ShowDialog();
            //Debug.WriteLine("This is the about button.");
        }

        private void closeBtn_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                this.Close();
            }
            catch (Exception ex)
            {
                Environment.Exit(0);
                Debug.WriteLine(ex.Message);
            }
        }

        private void Window_MouseDown(object sender, MouseButtonEventArgs e)
        {
            if (Mouse.LeftButton == MouseButtonState.Pressed)
            {
                this.DragMove();
            }
        }

        private void maximizeBtn_Click(object sender, RoutedEventArgs e)
        {
            //this.Close();
            if (this.WindowState == WindowState.Normal)
            {
                this.WindowState = WindowState.Maximized;
            }
            else if (this.WindowState == WindowState.Minimized)
            {
                this.WindowState = WindowState.Maximized;
            }
            else if (this.WindowState == WindowState.Maximized)
            {
                this.WindowState = WindowState.Normal;
            }
            //this.WindowState = WindowState.Maximized;

        }

        private void minimizeBtn_Click(object sender, RoutedEventArgs e)
        {
            if (this.WindowState == WindowState.Normal)
            {
                this.WindowState = WindowState.Minimized;
            }
            else if (this.WindowState == WindowState.Minimized)
            {
                this.WindowState = WindowState.Normal;
            }
            this.WindowState = WindowState.Minimized;
        }

        private void Window_SizeChanged(object sender, SizeChangedEventArgs e)
        {
            if (WindowState == WindowState.Normal)
            {
                this.maximizeBtn.ToolTip = "Maximize Window";
            }
            else if (WindowState == WindowState.Maximized)
            {
                this.maximizeBtn.ToolTip = "Restore Window";
            }
        }

        private void folderTxtBox_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            Settings.Default.filePathSetting = this.folderTxtBox.Text;
            Settings.Default.Save();
        }

        private void pdfBtn_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("this is the pdf button.");
            createPDF(MainWindow.currentItem);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) & ~WS_SYSMENU);
        }
    }
}


