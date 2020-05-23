using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Windows;
using System.Windows.Input;
using System.Windows.Navigation;
using Ookii.Dialogs.Wpf;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Data;
using PdfSharpCore;
using VetCV.HtmlRendererCore.PdfSharpCore;
using PdfSharpCore.Pdf;
using readClashReport.Properties;
using System.Drawing;

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
        public static string location { get; set; }
        private static string filename { get; set; }
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
                if (fdRun == true )
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

            return (filePathListTemp,valid);
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
                Debug.WriteLine(Path.Combine(this.folderTxtBox.Text, item.filename+".html"));
                this.webViewer.Navigate(Path.Combine(this.folderTxtBox.Text, item.filename + ".html"));
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
        private void emailBtn_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("this is the email button.");
            //string html = File.ReadAllText(@"C:\Users\dodzi\Desktop\Reports\ARCH (Ceilings, Roof Soffits) vs ARCH (00 Level).html");
            //PdfDocument pdf = PdfGenerator.GeneratePdf(html, PageSize.A4);
            //PdfDocument pdf = PdfGenerator.GeneratePdf(html, PageSize.A4);

            //pdf.Save(@"C:\Users\dodzi\Desktop\Reports\document.pdf");
            //excel.writeExcel.PdfSharpConvert(@"C:\Users\dodzi\Desktop\Reports\ARCH (Ceilings, Roof Soffits) vs ARCH (00 Level).html");
            excel.writeExcel.CreateMailItem("Sample Subject");
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

       


        //private void commonCommandBinding_CanExecute(object sender, CanExecuteRoutedEventArgs e)
        //{
        //    e.CanExecute = true;
        //}

        //private void wbSample_Navigating(object sender, System.Windows.Navigation.NavigatingCancelEventArgs e)
        //{
        //    Debug.WriteLine(e.Uri.OriginalString);
        //}
    }
}
