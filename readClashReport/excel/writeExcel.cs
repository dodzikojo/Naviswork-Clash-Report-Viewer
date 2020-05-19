using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.Diagnostics;

namespace readClashReport.excel
{
    class writeExcel
    {
        public static void writeExcelFile(string[,] data)
        {
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

            WriteToFile( workbook);


        }

        static void WriteToFile(IWorkbook workbook)
        {
            using (FileStream stream = new FileStream("outfile.xls", FileMode.Create, FileAccess.ReadWrite))
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
    }
}
