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
    class WriteToTable
    {
       
           public void GetTable(string SQLCode, string startCol, string endCol, int startRow, string destTabName)
            {

            //Gets Excel and gets Activeworkbook and worksheet
            Excel.Application oXL;
            Excel.Workbook oWB;
            oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Excel.Workbook)oXL.ActiveWorkbook;
            Excel.Sheets ExcelSheets = oWB.Worksheets;
            Excel.Worksheet excelWorksheet = (Excel.Worksheet)ExcelSheets.get_Item(destTabName);

            

            System.Data.DataTable dataTableSQL = new System.Data.DataTable();
            //write the sqlCode to dataTableSQL via adapter
            using (SqlConnection conn = new SqlConnection(@"Data Source=MRCEAUTO\SQLMRCEAUTO;Initial Catalog=CoMRC;Integrated Security = SSPI"))
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
               
                try
                {
                    adapter.SelectCommand = new SqlCommand(SQLCode, conn);
                    conn.Open();
                }
                catch
                {
                    Console.WriteLine(@"Unable to connect to: Data Source=MRCEAUTO\SQLMRCEAUTO;Initial Catalog=CoMRC;Integrated Security = SSPI . Contact Clark to get connected. Hit any key to exit.");
                    Console.ReadLine();
                    return;
                }
                
              //  adapter.FillSchema(dataTableSQL, SchemaType.Mapped);
               
                adapter.Fill(dataTableSQL);
                conn.Close();
            }

            


            int recordsNum = dataTableSQL.Rows.Count;
            int colNum = dataTableSQL.Columns.Count;

            //convert the datatableSQL to a strongly typed list:
            string[,] SQLs = new string[recordsNum + 1, colNum];
            for (int row = 0; row <= recordsNum - 1; ++row)
            {
                for (int col = 0; col <= colNum - 1; ++col)
                {


                    if (string.IsNullOrEmpty(dataTableSQL.Rows[row][col].ToString()))
                    {
                        SQLs[row, col] = "";
                    }
                    else
                    {

                        SQLs[row, col] = dataTableSQL.Rows[row][col].ToString().Replace(" ", String.Empty);
                    }

                }
            }





            //clear old data before pasting new
            var deleteRange = excelWorksheet.get_Range("A2","ZZ1000000");
            deleteRange.ClearContents();

            

            //paste new data
            var tablePaste = excelWorksheet.get_Range(startCol + startRow, endCol + (recordsNum + startRow));
            tablePaste.Value2 = SQLs;

            //var headerPaste = excelWorksheet.get_Range(startCol + (startRow-1), endCol + (startRow - 1));

            return;

            }





    }
}
