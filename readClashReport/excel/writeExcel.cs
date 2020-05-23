using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Diagnostics;
using Ookii.Dialogs.Wpf;
using VetCV.HtmlRendererCore.PdfSharpCore;
using PdfSharpCore.Pdf;
using PdfSharpCore;
using Microsoft.Office.Interop.Outlook;
using System.Threading.Tasks;

namespace readClashReport.excel
{
    class writeExcel
    {
        public static void writeExcelFile(string[,] data)
        {
            VistaFolderBrowserDialog folderBrowser = new VistaFolderBrowserDialog();
            folderBrowser.Description = "Select a location to save the excel file";
            folderBrowser.UseDescriptionForTitle = true;
            folderBrowser.ShowNewFolderButton = true;
            folderBrowser.SelectedPath = MainWindow.location;


            Nullable<bool> fdRun = folderBrowser.ShowDialog();

            if (fdRun == true)
            {
                string chosenPath = folderBrowser.SelectedPath;

                IWorkbook workbook = new HSSFWorkbook();
                ISheet sheet = workbook.CreateSheet("Clash Report");

                IFont font = workbook.CreateFont();
                //font.FontName = ("Bold");
                font.IsBold = true;
                font.FontHeightInPoints = 18;
                ICellStyle boldStyle = workbook.CreateCellStyle();
                boldStyle.SetFont(font);

                IRow titleRow = sheet.CreateRow(0);
                ICell title1_Cell = titleRow.CreateCell(0);
                ICell title2_Cell = titleRow.CreateCell(1);
                ICell title3_Cell = titleRow.CreateCell(2);
                ICell title4_Cell = titleRow.CreateCell(3);
                ICell title5_Cell = titleRow.CreateCell(4);
                ICell title6_Cell = titleRow.CreateCell(5);
                ICell title7_Cell = titleRow.CreateCell(6);
                ICell title8_Cell = titleRow.CreateCell(7);
                ICell title9_Cell = titleRow.CreateCell(8);

                title1_Cell.SetCellValue("Clash Test");
                title2_Cell.SetCellValue("Tolerance");
                title3_Cell.SetCellValue("Clashes");
                title4_Cell.SetCellValue("New Clashes");
                title5_Cell.SetCellValue("Active Clashes");
                title6_Cell.SetCellValue("Reviewed Clashes");
                title7_Cell.SetCellValue("Approved Clashes");
                title8_Cell.SetCellValue("Resolved Clashes");
                title9_Cell.SetCellValue("Type of Clash");

                IFont Standardfont = workbook.CreateFont();
                //font.FontName = ("Standard");
                Standardfont.IsBold = false;
                Standardfont.FontHeightInPoints = 9;
                ICellStyle standardStyle = workbook.CreateCellStyle();
                standardStyle.SetFont(Standardfont);

                try
                {
                    for (int rowNo = 0; rowNo < data.GetLength(0); rowNo++)
                    {
                        IRow row = sheet.CreateRow(rowNo + 1);
                        ICell cell_1 = row.CreateCell(0);
                        ICell cell_2 = row.CreateCell(1);
                        ICell cell_3 = row.CreateCell(2);
                        ICell cell_4 = row.CreateCell(3);
                        ICell cell_5 = row.CreateCell(4);
                        ICell cell_6 = row.CreateCell(5);
                        ICell cell_7 = row.CreateCell(6);
                        ICell cell_8 = row.CreateCell(7);
                        ICell cell_9 = row.CreateCell(8);
                        Debug.WriteLine($"Row is: {rowNo}");

                        for (int colNo = 0; colNo < data.GetLength(1); colNo++)
                        {
                            try
                            {
                                cell_1.SetCellValue(data[rowNo, 0]);
                                cell_2.SetCellValue(data[rowNo, 1]);
                                cell_3.SetCellValue(Int32.Parse(data[rowNo, 2]));
                                cell_4.SetCellValue(Int32.Parse(data[rowNo, 3]));
                                cell_5.SetCellValue(Int32.Parse(data[rowNo, 4]));
                                cell_6.SetCellValue(Int32.Parse(data[rowNo, 5]));
                                cell_7.SetCellValue(Int32.Parse(data[rowNo, 6]));
                                cell_8.SetCellValue(Int32.Parse(data[rowNo, 7]));
                                cell_9.SetCellValue(data[rowNo, 8]);
                            }
                            catch (System.Exception e)
                            {

                                Debug.WriteLine(e.Message);
                            }
                            

                        }
                    }

                }
                catch (System.Exception e)
                {

                    Debug.WriteLine(e.Message);
                }



                for (int i = 0; i <= data.GetLength(0); i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                DateTime thisDay = DateTime.Now;

                workbook.Close();
                string excelfileName = Path.Combine(chosenPath,"Clash Report Export_"+thisDay.ToString("yyyyMMdd")+".xls");
                WriteToFile(workbook, excelfileName);


                if (MainWindow.openExcelBool)
                {
                    var p = new Process();
                    p.StartInfo = new ProcessStartInfo(excelfileName)
                    {
                        UseShellExecute = true
                    };
                    p.Start();
                }


                
            }

        }

        static void WriteToFile(IWorkbook workbook, string savelocation)
        {
            //TODO: Warn users of potential block by antivirus and should be allowed.
            using (FileStream stream = new FileStream(savelocation, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(stream);
            }
        }

        static void writeColTitles(IWorkbook workbook)
        {
            IFont font = workbook.CreateFont();
            font.IsBold = true;
            font.FontHeightInPoints = 11;
            ICellStyle boldStyle = workbook.CreateCellStyle();
            boldStyle.SetFont(font);
        }

        public static Byte[] PdfSharpConvert(String html)
        {
            Byte[] res = null;
            using (MemoryStream ms = new MemoryStream())
            {
                var pdf = PdfGenerator.GeneratePdf(html, PageSize.A4);
                pdf.Save(ms);
                res = ms.ToArray();
            }
            return res;
        }

        public static async void CreateMailItem(string subject, string emailbody = "")
        {
            Microsoft.Office.Interop.Outlook.Application outApp = new Microsoft.Office.Interop.Outlook.Application();
            Microsoft.Office.Interop.Outlook.MailItem mailItem = (Microsoft.Office.Interop.Outlook.MailItem)outApp.CreateItem(Microsoft.Office.Interop.Outlook.OlItemType.olMailItem) as Microsoft.Office.Interop.Outlook.MailItem;
            mailItem.Subject = subject;
            mailItem.Body = emailbody;
            //using (StreamReader sr = File.OpenText(@"\\Data1\caddrive\CAD\6348-AWE-The_Hub\04-Production\01-Revit\01-WIP\01-BIM\10-Offline Studies\emaillist.txt"))
            //{
            //    string s = String.Empty;
            //    while ((s = sr.ReadLine()) != null)
            //    {
            //        mailItem.To = s;
            //    }
            //}
            await Task.Run(() =>
            {
                try
                {
                    mailItem.Display();
                }
                catch (System.Exception ex)
                {

                    Debug.WriteLine(ex.Message);
                }
            });


        }
    }
}
