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
        private static string filename { get; set; }
        public static List<htmlFiles> fileData = new List<htmlFiles>();
        public static List<string> filenamesList = new List<string>();
        public static List<string> clashesList = new List<string>();
        List<string> filePathList = new List<string>();

        /// <summary>
        /// Window Object
        /// </summary>
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                if (Properties.Appsettings.Default.filePathSetting.Length > 0)
                {
                    //this.folderTxtBox.Text = Properties.Appsettings.Default.filePathSetting;
                    getHTMLfiles(this.folderTxtBox.Text = Properties.Appsettings.Default.filePathSetting);
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
            string[,] tempData = new string[data.Count, 2];
            Debug.WriteLine(data.Count);
            try
            {
                foreach (string item in data)
                {
                    tempData[num, 0] = item;
                    try
                    {
                        tempData[num, 1] = clashesList[num];
                    }
                    catch (Exception ex1)
                    {

                        Debug.WriteLine(ex1.Message);
                    }
                    
                    num++;
                }

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
            Properties.Appsettings.Default.Save();
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
        }

        /// <summary>
        /// TODO: Save the .html file as PDF.
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void pdfBtn_Click(object sender, RoutedEventArgs e)
        {
            Debug.WriteLine("this is the save pdf button");
            excel.writeExcel.writeExcelFile(htmlFiles.data);
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
