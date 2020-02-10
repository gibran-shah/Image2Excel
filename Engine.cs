using System;
using System.Drawing;
//using GemBox.Spreadsheet; // https://www.gemboxsoftware.com/spreadsheet/examples/c-sharp-vb-net-excel-library/601
using Microsoft.Office.Interop.Excel;

namespace Image2Excel {
    class Engine {
        public static void go(string imageFilename, string excelFilename = null) {
            //Image img = Image.FromFile(imageFilename);
            Bitmap btm = (Bitmap)Bitmap.FromFile(imageFilename, false);
            Image img = Image.FromFile(imageFilename, false);

            if (btm != null) {
                Color[][] colorArray = new Color[btm.Width][];
                for (int x = 0; x < btm.Width; x++) {
                    colorArray[x] = new Color[btm.Height];
                    for (int y = 0; y < btm.Height; y++) {
                        colorArray[x][y] = btm.GetPixel(x, y);
                    }
                }

// https://www.c-sharpcorner.com/UploadFile/bd6c67/how-to-create-excel-file-using-C-Sharp/

                var excel = new Microsoft.Office.Interop.Excel.Application();
                var workbook = excel.Workbooks.Add(Type.Missing);
                var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.ActiveSheet;
                worksheet.Name = "sheet1";

                worksheet.Cells[1,1] = "top left";
                worksheet.Cells[1,2] = "top right";
                worksheet.Cells[2,1] = "bottom left";
                worksheet.Cells[2,2] = "bottom right";

                workbook.SaveAs("temp.xlsx");
                workbook.Close();
                excel.Quit();
/*
                SpreadsheetInfo.SetLicense("FREE-LIMITED-KEY");
                var workbook = new ExcelFile();
                var worksheet = workbook.Worksheets.Add("page 1");

                // worksheet.Cells[0,0].Value = "top left";
                // worksheet.Cells[0,1].Value = "top right";
                // worksheet.Cells[1,0].Value = "bottom left";
                // worksheet.Cells[1,1].Value = "bottom right";


                // Console.WriteLine("DefaultRowHeight: " + worksheet.DefaultRowHeight);
                // Console.WriteLine("DefaultColumnWidth: " + worksheet.DefaultColumnWidth);
                
                worksheet.DefaultRowHeight = 250;
                worksheet.DefaultColumnWidth = 500;

                // worksheet.Cells[0,0].Style.FillPattern.SetPattern(FillPatternStyle.Solid, Color.Red, Color.Red);
                // worksheet.Cells[1,0].Style.FillPattern.SetPattern(FillPatternStyle.Solid, Color.Blue, Color.Blue);
                // worksheet.Cells[0,1].Style.FillPattern.SetPattern(FillPatternStyle.Solid, Color.Green, Color.Green);
                // worksheet.Cells[1,1].Style.FillPattern.SetPattern(FillPatternStyle.Solid, Color.Yellow, Color.Yellow);

                for (int x = 0; x < btm.Width; x++) {
                    colorArray[x] = new Color[btm.Height];
                    for (int y = 0; y < btm.Height; y++) {
                        worksheet.Cells[x,y].Style.FillPattern.SetPattern(FillPatternStyle.Solid, colorArray[x][y], colorArray[x][y]);
                    }
                }

                workbook.Save("temp.xlsx");
*/

                // Bitmap newBtm = null;
                
                // try {
                //     newBtm = (Bitmap)Bitmap.FromFile("newfile.btm", false);
                // } catch (Exception e) {
                //     Console.WriteLine("Error: " + e.Message);
                //     return;
                // }

                // for (int x = 0; x < btm.Width; x++) {
                //     for (int y = 0; y < btm.Height; y++) {
                //         newBtm.SetPixel(x, y, colorArray[x][y]);
                //     }
                // }

                // newBtm.Save("newfile.btm");
            }
        } 
    }
}