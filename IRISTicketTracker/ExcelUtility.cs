using System;
using System.Collections.Generic;
using System.Text;
using System.Data;
using System.IO;
using OfficeOpenXml;
using System.Linq;

namespace IRISTicketTracker
{
    public class ExcelUtility
    {
        public void ReadData()
        {
            var fileInfo = new FileInfo(@"D:\CR data v1.xlsx");
            using (var package = new ExcelPackage(fileInfo))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets[1];
                int colCount = worksheet.Dimension.End.Column;  //get Column Count
                int rowCount = worksheet.Dimension.End.Row;   //get row count
                DataTable dt = new DataTable();

                for (int row = 1; row <= rowCount; row++)
                {
                    var dr = dt.NewRow();

                    for (int col = 1; col <= colCount; col++)
                    {
                        if (row == 1) // going to define the dt header
                        {
                            var dc = new DataColumn(worksheet.Cells[row, col].Value.ToString().Trim());
                            dt.Columns.Add(dc);
                        }
                        else // going to add data to dt
                        {
                            dr[col - 1] = worksheet.Cells[row, col].Value.ToString().Trim();
                        }

                        Console.WriteLine(" Row:" + row + " column:" + col + " Value:" + worksheet.Cells[row, col].Value.ToString().Trim());
                    }

                    if (row != 1)
                        dt.Rows.Add(dr);
                }
            }
        }
    }
}
