using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;
using System.Data;
using System.ComponentModel;
using System.Data.SqlClient;
using System.Collections;
using System.Data.OleDb;
using System.Text.RegularExpressions;

namespace RunSqlFromExcel
{
    class ConvertToNum
    {

        public void ConvertMethod(string numCol, string destTabName)
        {

            //Gets Excel and gets Activeworkbook and worksheet
            Excel.Application oXL;
            Excel.Workbook oWB;
            oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Excel.Workbook)oXL.ActiveWorkbook;
            Excel.Sheets ExcelSheets = oWB.Worksheets;
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)ExcelSheets.get_Item(destTabName);





            //convert number range to number 
            var rangeOfValues = excelWorksheet.get_Range(numCol);
            rangeOfValues.Cells.NumberFormat = "#,##0.00";
            rangeOfValues.Value = rangeOfValues.Value;


            return;
        }



    }
}
