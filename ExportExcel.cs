using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;

namespace C3D_table_export
{
    public class ExportExcel
    {
        public static void ExportToExcel(List<TableData> list)
        {

            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                MessageBox.Show("Excel is not properly installed!!");
                return;
            }

            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            for (int i = 0; i < list.Count; i++)
            {
                int row = int.Parse(list[i].Line);
                List<string> values = list[i].Values;
                int index = 0;

                foreach (string x in values)
                {
                    index++;
                    xlWorkSheet.Cells[row + 1, index] = values[index - 1];
                }
            }
            //xlWorkSheet.Cells[1, 1] = list[0].Values[0];
            //xlWorkSheet.Cells[1, 2] = list[0].Values[1];
            //xlWorkSheet.Cells[1, 3] = list[0].Values[2];

            xlWorkBook.SaveAs("c:\\temp\\csharp-Excel.xls", Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();


            MessageBox.Show("Excel file created , you can find the file csharp-Excel.xls");

        }
    }
}


