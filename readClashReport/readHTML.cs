using HtmlAgilityPack;
using System;
using System.Diagnostics;
using System.Collections.Generic;
using System.Text;
using System.IO;

namespace readClashReport
{
    class readHTML
    {
        private static int clashesCounter { get; set; }
        private static int toleranceCounter { get; set; }
        private static int newcounter { get; set; }
        private static int reviewedCounter { get; set; }
        private static int approvedCounter { get; set; }
        private static int activeCounter { get; set; }
        private static int resolvedCounter { get; set; }
        private static int typeCounter { get; set; }

        public static htmlFiles fileData { get; set; }

        

        //private static Stopwatch watch;

         
        public static void readHTMLData(string html)
        {
            
            var htmlDoc = new HtmlDocument();
            htmlDoc.Load(html);
            //htmlDoc.LoadHtml(html);


            var htmlNodes = htmlDoc.DocumentNode.SelectNodes("//body/table");
           // MainWindow.fileData.Add(new htmlFiles() { filename = html });

            toleranceCounter = 0;
            clashesCounter = 0;
            newcounter = 0;
            reviewedCounter = 0;
            approvedCounter = 0;
            resolvedCounter = 0;
            typeCounter = 0;
            activeCounter = 0;

            string filenameTemp = string.Empty;
            string clashesCounttemp = string.Empty;
            string activeTemp = string.Empty;
            string newTemp = string.Empty;
            string typeTemp = string.Empty;
            string toleranceTemp = string.Empty;
            string resolvedTemp = string.Empty;

            foreach (var node in htmlNodes)
            {
                if (node.HasClass("testSummaryTable"))
                {
                    var abc = node.Descendants();

                    foreach (var nodes in abc)
                    {

                        if (nodes.HasClass("headerCell"))
                        {
                            toleranceCounter++;
                            if (nodes.InnerText.ToString() == "Tolerance")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{toleranceCounter}]");
                                //Debug.WriteLine($"Tolerance is: {nodess.InnerText.ToString()}");
                                
                            }
                        }


                        if (nodes.HasClass("headerCell"))
                        {
                            clashesCounter++;
                            if (nodes.InnerText.ToString() == "Clashes")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{clashesCounter}]");
                                //Debug.WriteLine($"Total number of clashes are: {nodess.InnerText.ToString()}");
                                clashesCounttemp = nodess.InnerText.ToString();
                                //filenameTemp = html;
                                //MainWindow.fileData.Add(new htmlFiles() { clashes = nodess.InnerText.ToString(), filename = html });
                            }
                        }

                        if (nodes.HasClass("headerCell"))
                        {
                            newcounter++;
                            if (nodes.InnerText.ToString() == "New")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{newcounter}]");
                                //MainWindow.fileData.Add(new htmlFiles() { newClashes = nodess.InnerText.ToString()});
                                newTemp = nodess.InnerText.ToString();
                                //Debug.WriteLine($"Number of new clashes: {nodess.InnerText.ToString()}");
                            }
                        }

                        if (nodes.HasClass("headerCell"))
                        {
                            activeCounter++;
                            if (nodes.InnerText.ToString() == "Active")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{activeCounter}]");
                                activeTemp = nodess.InnerText.ToString();
                                //Debug.WriteLine($"Number of active clashes: {nodess.InnerText.ToString()}");
                            }
                        }

                        if (nodes.HasClass("headerCell"))
                        {
                            reviewedCounter++;
                            if (nodes.InnerText.ToString() == "Reviewed")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{reviewedCounter}]");
                                
                                //Debug.WriteLine($"Number of reviewed clashes: {nodess.InnerText.ToString()}");

                            }
                        }

                        if (nodes.HasClass("headerCell"))
                        {
                            approvedCounter++;
                            if (nodes.InnerText.ToString() == "Approved")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{approvedCounter}]");
                                //Debug.WriteLine($"Number of approved clashes: {nodess.InnerText.ToString()}");
                            }
                        }

                        if (nodes.HasClass("headerCell"))
                        {
                            resolvedCounter++;
                            if (nodes.InnerText.ToString() == "Resolved")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{resolvedCounter}]");
                                //Debug.WriteLine($"Number of resolved clashes: {nodess.InnerText.ToString()}");
                            }
                        }

                        if (nodes.HasClass("headerCell"))
                        {
                            typeCounter++;
                            if (nodes.InnerText.ToString() == "Type")
                            {
                                var nodess = htmlDoc.DocumentNode.SelectSingleNode($"(//body/table[2]/tr[@class='contentRow'])/td[{typeCounter}]");
                                //Debug.WriteLine($"Type of clash performed: {nodess.InnerText.ToString()}");
                                //Debug.WriteLine("\n");
                            }
                        }

                    }
                    string file = Path.GetFileNameWithoutExtension(html);
                    MainWindow.fileData.Add(new htmlFiles() { clashes = clashesCounttemp, filename = file, newClashes = newTemp });

                    try
                    {
                        MainWindow.filenamesList.Add(file);
                        MainWindow.clashesList.Add(clashesCounttemp);
                    }
                    catch (Exception e)
                    {

                        Debug.WriteLine(e.Message);
                    }
                    
                }
            }




            //int num1 = 2;


        }
    }
}
