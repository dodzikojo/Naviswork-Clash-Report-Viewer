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

                title1_Cell.SetCellValue("Clash Test");
                title2_Cell.SetCellValue("Number of Clashes");

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

                        for (int colNo = 0; colNo < data.GetLength(1); colNo++)
                        {
                            cell_1.SetCellValue(data[rowNo, 0]);
                            cell_2.SetCellValue(Int32.Parse(data[rowNo, 1]));

                        }
                    }

                }
                catch (Exception e)
                {

                    Debug.WriteLine(e.Message);
                }



                for (int i = 0; i <= data.GetLength(0); i++)
                {
                    sheet.AutoSizeColumn(i);
                }

                DateTime thisDay = DateTime.Now;

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
            using (FileStream stream = new FileStream(savelocation, FileMode.Create, FileAccess.ReadWrite))
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
    }
}
