using HtmlAgilityPack;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Text;

namespace readClashReport.reader_classes
{
    class readClashData
    {
        public static ReadClashData clash { get; set; }
        public static int overallCounter { get; set; }

        /// <summary>
        /// Reads each clash data from Navisworks HTML report.
        /// </summary>
        /// <param name="file"></param>
        public static void ReadHTML_ClashData(string file)
        {
            int clashNameCounter = 0;
            int statusCounter = 0;
            int descripCounter = 0;
            int dateFoundCounter = 0;
            int clashPointCounter = 0;
            int distanceCounter = 0;
            int layer_1_Counter = 0;
            int layer_2_Counter = 0;
            int itemID_1_Counter = 0;
            int itemName_1_Counter = 0;
            int itemSource_1_Counter = 0;
            int itemType_1_Counter = 0;
            int itemID_2_Counter = 0;
            int itemName_2_Counter = 0;
            int itemSource_2_Counter = 0;
            int itemType_2_Counter = 0;

            HtmlDocument htmlDoc = new HtmlDocument();

            htmlDoc.Load(@file);

            HtmlNodeCollection headerNodes = htmlDoc.DocumentNode.SelectNodes("//body/table[3]/tr[2]");

            int num = 0;
            foreach (HtmlNode node in headerNodes)
            {

                var collection = node.ChildNodes;
                foreach (var item in collection)
                {

                    if (!String.IsNullOrEmpty(item.InnerText))
                    {
                        num++;
                        if (item.InnerText.ToString().ToLower() == "clash name")
                        {
                            clashNameCounter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "clash point")
                        {
                            clashPointCounter = num;
                        }
                        if (item.InnerText.ToString().ToLower() == "distance")
                        {
                            distanceCounter = num;

                        }
                        if (item.InnerText.ToString().ToLower() == "item id" && itemID_1_Counter == 0 && itemID_2_Counter == 0)
                        {
                            itemID_1_Counter = num;
                        }
                        else if (item.InnerText.ToString().ToLower() == "item id" && itemID_1_Counter != 0 && itemID_2_Counter == 0)
                        {
                            itemID_2_Counter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "item name" && itemName_1_Counter == 0)
                        {
                            itemName_1_Counter = num;
                        }
                        else if (item.InnerText.ToString().ToLower() == "item name" && itemName_1_Counter != 0 && itemName_2_Counter == 0)
                        {
                            itemName_2_Counter = num;
                        }
                        if (item.InnerText.ToString().ToLower() == "layer" && layer_1_Counter == 0)
                        {
                            layer_1_Counter = num;
                        }
                        else if (item.InnerText.ToString().ToLower() == "layer" && layer_1_Counter != 0 && layer_2_Counter == 0)
                        {
                            layer_2_Counter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "item source file" && itemSource_1_Counter == 0)
                        {
                            itemSource_1_Counter = num;
                        }

                        else if (item.InnerText.ToString().ToLower() == "item source file" && itemSource_1_Counter != 0 && itemSource_2_Counter == 0)
                        {
                            itemSource_2_Counter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "item type" && itemType_1_Counter == 0)
                        {
                            itemType_1_Counter = num;
                        }

                        else if (item.InnerText.ToString().ToLower() == "item type" && itemType_1_Counter != 0 && itemType_2_Counter == 0)
                        {
                            itemType_2_Counter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "status")
                        {
                            statusCounter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "description")
                        {
                            descripCounter = num;
                        }

                        if (item.InnerText.ToString().ToLower() == "date found")
                        {
                            dateFoundCounter = num;
                        }


                    }

                }

            }

            HtmlNodeCollection contentNodes = htmlDoc.DocumentNode.SelectNodes("//body/table[3]/tr");

            int TableCount = 0;
            foreach (HtmlNode node in contentNodes)
            {

                var collection = node.ChildNodes;
                if (TableCount > 1)
                {
                    int num1 = 0;
                    clash = new readClashReport.ReadClashData();
                    foreach (var item in collection)
                    {

                        num1++;
                        if (num1 == clashNameCounter)
                        {
                            string value = item.InnerText;
                            clash.clashName = value;
                        }
                        if (num1 == clashPointCounter)
                        {
                            string IDText = item.InnerText.Trim();
                            string finalValue = String.Empty;
                            foreach (char character in IDText)
                            {
                                if (char.IsDigit(character) || character == ',' || character == '.')
                                {
                                    finalValue += character;
                                }
                            }
                            clash.clashPoint = finalValue;
                        }
                        if (num1 == itemID_1_Counter)
                        {

                            string IDText = item.InnerText.Trim();
                            string finalValue = String.Empty;
                            foreach (char character in IDText)
                            {
                                if (char.IsDigit(character))
                                {
                                    finalValue += character;
                                }
                            }
                            clash.itemID_1 = finalValue;
                        }
                        if (num1 == layer_1_Counter)
                        {
                            string value = item.InnerText.Trim();
                            clash.layer_1 = value;
                        }

                        if (num1 == itemID_2_Counter)
                        {
                            string IDText = item.InnerText.Trim();
                            string finalValue = String.Empty;
                            foreach (char character in IDText)
                            {
                                if (char.IsDigit(character))
                                {
                                    finalValue += character;
                                }
                            }
                            clash.itemID_2 = finalValue;

                        }
                        if (num1 == layer_2_Counter)
                        {
                            string value = item.InnerText;
                            clash.layer_2 = value;

                        }
                        if (num1 == distanceCounter)
                        {
                            string value = item.InnerText;
                            clash.distance = value;
                        }
                        if (num1 == statusCounter)
                        {
                            string value = item.InnerText;
                            clash.status = value;
                        }


                        if (num1 == descripCounter)
                        {
                            string value = item.InnerText;
                            clash.description = value;

                        }
                        if (num1 == itemName_1_Counter)
                        {
                            string value = item.InnerText;
                            clash.itemName_1 = value;

                        }
                        if (num1 == itemName_2_Counter)
                        {
                            string value = item.InnerText;
                            clash.itemName_2 = value;

                        }
                        if (num1 == itemSource_1_Counter)
                        {
                            string value = item.InnerText;
                            clash.itemSource_1 = value;
                        }
                        if (num1 == itemSource_2_Counter)
                        {
                            string value = item.InnerText;
                            clash.itemSource_2 = value;

                        }
                        if (num1 == itemType_1_Counter)
                        {
                            string value = item.InnerText;
                            clash.itemType_1 = value;
                        }
                        if (num1 == itemType_2_Counter)
                        {
                            string value = item.InnerText;
                            clash.itemType_2 = value;

                        }
                        if (num1 == dateFoundCounter)
                        {
                            string date = item.InnerText.Trim();
                            DateTime datetime = Convert.ToDateTime(date);
                            clash.date = datetime.ToString("F");

                        }

                    }
                    MainWindow.allClashData.Add(readClashData.clash);
                    overallCounter++;
                }
                
                TableCount++;
            }
        }
    }
}
