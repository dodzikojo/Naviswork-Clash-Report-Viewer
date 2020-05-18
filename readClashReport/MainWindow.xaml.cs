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
    }
    
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        private static string filename { get; set; }
        public static List<htmlFiles> fileData = new List<htmlFiles>();
        string[] filePaths;
        List<string> filePathList = new List<string>();
        public MainWindow()
        {
            InitializeComponent();
            try
            {
                if (Properties.Appsettings.Default.filePathSetting.Length > 0)
                {
                    this.folderTxtBox.Text = Properties.Appsettings.Default.filePathSetting;
                    filePaths = Directory.GetFiles(this.folderTxtBox.Text, "*.html");
                    this.countLabel.Content = filePaths.Length.ToString();
                    if (isValid(filePaths))
                    {
                        DisplayData();
                    }
                    
                }
            }
            catch (Exception e)
            {
                Debug.WriteLine(e.Message);
            }
        }

        private void browseBtn_Click(object sender, RoutedEventArgs e)
        {
            bool valid = false;
            VistaFolderBrowserDialog folderBrowser = new VistaFolderBrowserDialog();
            Nullable<bool> fdRun = folderBrowser.ShowDialog();
            try
            {
                if (fdRun == true )
                {
                    string[] tempFilepaths;
                    this.folderTxtBox.Text = folderBrowser.SelectedPath.ToString();
                    tempFilepaths = Directory.GetFiles(folderBrowser.SelectedPath.ToString(), "*.html");
                    if (tempFilepaths.Length > 0)
                    {
                        fileData.Clear();
                        valid = isValid(tempFilepaths);
                        this.countLabel.Content = filePathList.Count.ToString();
                    }
                    if (valid)
                    {
                        fileData.Clear();
                        DisplayData();
                        Properties.Appsettings.Default.filePathSetting = folderBrowser.SelectedPath.ToString();
                    }
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
        public bool isValid(string[] tempFilepaths)
        {
            bool valid;
            filePathList.Clear();
            if (tempFilepaths.Length > 0)
            {
                foreach (string item in tempFilepaths)
                {
                    if (Path.GetFileNameWithoutExtension(item).ToLower().Contains("vs") && Path.GetExtension(item).ToString() == ".html")
                    {
                        filePathList.Add(item);
                    }
                }
                valid = true;
            }
            else
            {
                valid = false;
                MessageBox.Show("No valid files found.","Error",MessageBoxButton.OK,MessageBoxImage.Error,MessageBoxResult.OK);
            }
            return valid;
        }

        /// <summary>
        /// Displays .html files in the listview panel.
        /// </summary>
        public async void DisplayData()
        {
            filesListView.ItemsSource = null;
            filesListView.Items.Clear();
            fileData.Clear();
            await Task.Run(() =>
            {
                try
                {
                    foreach (string html in filePathList)
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
                        this.webViewer.Navigate(filePathList[0]);

                    });

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
