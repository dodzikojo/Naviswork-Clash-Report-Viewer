using Newtonsoft.Json;
using Ookii.Dialogs.Wpf;
using PuppeteerSharp;
using readClashReport.Information.UI;
using readClashReport.Properties;
using readClashReport.reader_classes;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Diagnostics;
using System.IO;
using System.Runtime.InteropServices;
using System.Threading;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Interop;

namespace readClashReport
{
   

    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private const int GWL_STYLE = -16;
        private const int WS_SYSMENU = 0x80000;

        public static List<ReadClashData> allClashData = new List<ReadClashData>();
        public static string location { get; set; }
        private static string filename { get; set; }
        public static string currentItem { get; set; }
        public static List<ReadHtmlFiles> fileData = new List<ReadHtmlFiles>();
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
        public static string[,] data { get; set; }
        //public static string sample;



        /// <summary>
        /// Window Object
        /// </summary>
        public MainWindow()
        {
            //InitializeComponent();

            this.Dispatcher.Invoke(() =>
            {
                InitializeComponent();

                
                
                //Set window location
                try
                {
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
                }
                catch (Exception e)
                {
                    Debug.WriteLine(e.Message);
                }

            });

            

            try

            {
                if (Properties.Settings.Default.filePathSetting.Length > 0)
                {
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
            string[] fileArray = Directory.GetFiles(path, "*.html",SearchOption.AllDirectories);
            location = path;
            var results = isValid(fileArray);
            

            if (results.Item2)
            {
                this.folderTxtBox.Text = path;
                //htmlFiles.data = HTMLdata2DArr(results.Item1);
                
                filePathList = results.Item1;
                DisplayData(filePathList);
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
            //filePathList.Clear();
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
            toleranceList = new List<string>();
            filenamesList = new List<string>();
            clashesList = new List<string>();
            newList = new List<string>();
            activeList = new List<string>();
            reviewedList = new List<string>();
            approvedList = new List<string>();
            resolvedList = new List<string>();
            typeList = new List<string>();

            try
            {
                filesListView.ItemsSource = null;
                filesListView.Items.Clear();
                //fileData.Clear();
                //htmlFiles.data = null;
            }
            catch (Exception e)
            {

                Debug.WriteLine(e.Message);
            }
            this.InfoTipText.Text = "Reading HTML documents...";
            this.loadIcon.Spin = true;
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
                        //string selection = htmlFiles.data[selectedIndex, 0];
                        //this.webViewer.Navigate(selection);
                        this.FilePathText.Text = filepathlist[0];
                        currentItem = filepathlist[0];
                        location = filepathlist[0];

                    });
                    data = HTMLdata2DArr(filenamesList);
                    

                }
                catch (Exception ex)
                {
                    Debug.WriteLine(ex.Message);
                }
            });
            this.InfoTipText.Text = "Load completed";
            this.loadIcon.Spin = false;
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
                Console.WriteLine("Press any key to continue...");
                Console.ReadLine();
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
                return ((item as ReadHtmlFiles).filename.IndexOf(txtFilter.Text, StringComparison.OrdinalIgnoreCase) >= 0);
        }

        private void txtFilter_TextChanged(object sender, System.Windows.Controls.TextChangedEventArgs e)
        {
            CollectionViewSource.GetDefaultView(filesListView.ItemsSource).Refresh();
        }


        private void filesListView_MouseDoubleClick(object sender, MouseButtonEventArgs e)
        {

            int selectedIndex = this.filesListView.SelectedIndex;
            if (selectedIndex >= 0)
            {
                if (this.txtFilter.Text.Length > 0)
                {
                    string filtertext = this.txtFilter.Text;
                    this.txtFilter.Clear();
                   // Debug.WriteLine(this.filesListView.SelectedIndex);

                    string selection = data[this.filesListView.SelectedIndex, 0];
                    this.webViewer.Navigate(selection);
                    this.FilePathText.Text = selection;
                    this.txtFilter.Text = filtertext;
                }
                else
                {
                    string selection = data[selectedIndex, 0];
                    this.webViewer.Navigate(selection);
                    this.FilePathText.Text = selection;
                    location = selection;
                }
                

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
        /// TODO: Figure out why Creatinig PDF does not work in production.
        /// </summary>
        /// <param name="file"></param>
        static async Task createPDF(string file)
        {
            try
            {
                await new BrowserFetcher().DownloadAsync(BrowserFetcher.DefaultRevision);
                var browser = await Puppeteer.LaunchAsync(new LaunchOptions
                {
                    Headless = true
                });
                var page = await browser.NewPageAsync();
                await page.SetViewportAsync(new ViewPortOptions { Width = 50000, Height = 50 });

                NavigationOptions navOpts = new NavigationOptions();
                navOpts.Timeout = 0;
                await page.GoToAsync(file, navOpts);

                PdfOptions options = new PdfOptions();
                options.Width = 2000;
                options.Height = 1200;
                options.MarginOptions.Left = "100";
                options.MarginOptions.Right = "100";
                options.MarginOptions.Top = "50";
                options.MarginOptions.Bottom = "50";

                await page.PdfAsync(Path.Combine(Path.GetDirectoryName(location), Path.GetFileNameWithoutExtension(file) + ".pdf"), options);

                
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
           

        }

        /// <summary>
        /// TODO: Save the .html file as PDF.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void excelBtn_Click(object sender, RoutedEventArgs e)
        {
            //Debug.WriteLine("this is the save pdf button");
            this.InfoTipText.Text = "Starting Excel export";
            excel.writeExcel.writeExcelFile(data);
            this.InfoTipText.Text = "";
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
            try
            {
                var t = createPDF(MainWindow.currentItem);

                if (!t.IsCompleted)
                {
                    this.InfoTipText.Text = "PDF export may take longer to complete in some cases, active internet connection may be required.";
                }
                else
                {
                    this.InfoTipText.Text = "Completed";
                }
                
            }
            catch (Exception ex)
            {

                Debug.Assert(true, ex.Message, "Info");
            }

        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            var hwnd = new WindowInteropHelper(this).Handle;
            SetWindowLong(hwnd, GWL_STYLE, GetWindowLong(hwnd, GWL_STYLE) & ~WS_SYSMENU);
        }


       

        private void Sort(string sortBy, ListSortDirection direction)
        {
            ICollectionView dataView =
              CollectionViewSource.GetDefaultView(this.filesListView.ItemsSource);

            dataView.SortDescriptions.Clear();
            SortDescription sd = new SortDescription(sortBy, direction);
            dataView.SortDescriptions.Add(sd);
            dataView.Refresh();
        }

        #region Sort Listview
        //GridViewColumnHeader _lastHeaderClicked = null;
        //ListSortDirection _lastDirection = ListSortDirection.Ascending;


        //void GridViewColumnHeaderClickedHandler(object sender,
        //                                                RoutedEventArgs e)
        //{
        //    var headerClicked = e.OriginalSource as GridViewColumnHeader;
        //    ListSortDirection direction;

        //    if (headerClicked != null)
        //    {
        //        if (headerClicked.Role != GridViewColumnHeaderRole.Padding)
        //        {
        //            if (headerClicked != _lastHeaderClicked)
        //            {
        //                direction = ListSortDirection.Ascending;
        //            }
        //            else
        //            {
        //                if (_lastDirection == ListSortDirection.Ascending)
        //                {
        //                    direction = ListSortDirection.Descending;
        //                }
        //                else
        //                {
        //                    direction = ListSortDirection.Ascending;
        //                }
        //            }

        //            var columnBinding = headerClicked.Column.DisplayMemberBinding as Binding;
        //            var sortBy = columnBinding?.Path.Path ?? headerClicked.Column.Header as string;

        //            //string header = headerClicked.Column.Header as string;
        //            Sort(sortBy, direction);

        //            //Sort(sortBy, direction);

        //            if (direction == ListSortDirection.Ascending)
        //            {
        //                headerClicked.Column.HeaderTemplate =
        //                  Resources["HeaderTemplateArrowUp"] as DataTemplate;
        //            }
        //            else
        //            {
        //                headerClicked.Column.HeaderTemplate =
        //                  Resources["HeaderTemplateArrowDown"] as DataTemplate;
        //            }

        //            // Remove arrow from previously sorted header
        //            if (_lastHeaderClicked != null && _lastHeaderClicked != headerClicked)
        //            {
        //                _lastHeaderClicked.Column.HeaderTemplate = null;
        //            }

        //            _lastHeaderClicked = headerClicked;
        //            _lastDirection = direction;
        //        }
        //    }
        //}
        #endregion

        private async void jsonBtn_Click(object sender, RoutedEventArgs e)
        {
            this.loadIcon.Spin = true;

            allClashData.Clear();
            Stopwatch stopwatch = new Stopwatch();
            stopwatch.Start();

            await Task.Run(() =>
            {
                foreach (string item in filePathList)
                {
                    Debug.WriteLine(item);
                    readClashData.ReadHTML_ClashData(item);
                }
            });

            TimeSpan ts = stopwatch.Elapsed;

            Debug.WriteLine(ts.TotalSeconds);

            Debug.WriteLine($"Total number of clashes in the folder is: {readClashData.overallCounter}.");

            string json = JsonConvert.SerializeObject(allClashData, Formatting.Indented);

            // serialize JSON to a string and then write string to a file
            File.WriteAllText(@"C:\Users\Dodzi\Desktop\HTML Reports\list.json", json);


            this.loadIcon.Spin = false;

        }
    }
    
}


