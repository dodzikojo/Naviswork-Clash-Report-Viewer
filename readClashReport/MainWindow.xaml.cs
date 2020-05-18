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

                    //List<htmlFiles> fileData = new List<htmlFiles>();
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

       

        private void webBrowser_Nav(object sender, NavigatingCancelEventArgs e)
        {
            //if (!willNavigate)
            //{
            //    willNavigate = true;
            //    return;
            //}

            e.Cancel = false;

            //Process.Start(e.Uri.ToString());

            try
            {
                var startInfo = new ProcessStartInfo
                {
                    FileName = e.Uri.ToString()
                };

                Process.Start(e.Uri.OriginalString); 
            }
            catch (Exception ex2)
            {

                Debug.WriteLine(ex2.Message);
            }
            
            //Debug.WriteLine(e.Uri.ToString());
            //Debug.WriteLine(e.Uri.OriginalString);
        }

        private void webBrowser_Navigated(object sender, NavigationEventArgs e)
        {
            Debug.WriteLine(e.Uri.ToString());
        }

        private void webBrowser_LoadCompleted(object sender, NavigationEventArgs e)
        {
           
        }

        private  void browseBtn_Click(object sender, RoutedEventArgs e)
        {
            bool valid = false;
            VistaFolderBrowserDialog abcd = new VistaFolderBrowserDialog();
            Nullable<bool> sample = abcd.ShowDialog();
            try
            {
                //bool valid;
                if (sample == true )
                {
                    string[] tempFilepaths;
                    this.folderTxtBox.Text = abcd.SelectedPath.ToString();
                    tempFilepaths = Directory.GetFiles(abcd.SelectedPath.ToString(), "*.html");
                    if (tempFilepaths.Length > 0)
                    {
                        fileData.Clear();
                        valid = isValid(tempFilepaths);
                        
                        


                        //List<htmlFiles> fileData = new List<htmlFiles>();
                        this.countLabel.Content = filePathList.Count.ToString();
                        
                    }

                    if (valid)
                    {
                        fileData.Clear();
                        DisplayData();
                        Properties.Appsettings.Default.filePathSetting = abcd.SelectedPath.ToString();
                    }
                    


                }
            }
            catch (Exception ex1)
            {

                Debug.WriteLine(ex1.Message);
            }
           

            
           

        }

        public bool isValid(string[] tempFilepaths)
        {
            bool valid;
            filePathList.Clear();
            if (tempFilepaths.Length > 0)
            {
                foreach (string item in tempFilepaths)
                {
                    if (Path.GetFileNameWithoutExtension(item).ToLower().Contains("vs"))
                    {
                        filePathList.Add(item);
                    }
                }
                valid = true;
            }
            else
            {
                valid = false;
                MessageBox.Show("Error","No valid files found.");
            }
            return valid;
        }

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
                        //fileData.Add(new htmlFiles() { filename = "lajlfaa" });

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
            //System.Collections.IList<string> abcd =  this.filesListView.SelectedItems;

            foreach (htmlFiles item in this.filesListView.SelectedItems)
            {
                Debug.WriteLine(Path.Combine(this.folderTxtBox.Text, item.filename+".html"));
                this.webViewer.Navigate(Path.Combine(this.folderTxtBox.Text, item.filename + ".html"));
            }

            //if (this.filesListView.SelectedIndex < 0)
            //{

            //}
            //else
            //{
            //    this.webViewer.Navigate(filePaths[this.filesListView.SelectedIndex]);
            //    //Path.Combine(this.folderTxtBox.Text, item.filename);
            //}
            

            //Debug.WriteLine(filesListView.SelectedIndex.ToString());
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            Properties.Appsettings.Default.Save();
            e.Cancel = true;
            //this.Visibility = Visibility.Visible;
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

        //private void wbSample_Navigating(object sender, System.Windows.Navigation.NavigatingCancelEventArgs e)
        //{
        //    Debug.WriteLine(e.Uri.OriginalString);
        //}
    }
}
