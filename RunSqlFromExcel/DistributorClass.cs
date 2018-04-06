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
    class DistributorClass
    {

        public void DistributorMethod(string begPeriodRef, string endPeriodRef, string TSQLType, string startCol, string endCol, int startRow, string numberRange, string siteID, string reportingPeriod, string destTabName, string report)
        {



        //get sql code for query
        RunSqlFromExcel.getSQLClass getSQL = new RunSqlFromExcel.getSQLClass();

            //beg period date, end period date, SQL Identifier
            string SQLCode = null;
           
               SQLCode = getSQL.getSQLMethod(begPeriodRef, endPeriodRef, TSQLType, siteID, reportingPeriod, report);
          

            //write the code to a table
           
                RunSqlFromExcel.WriteToTable wtt = new RunSqlFromExcel.WriteToTable();
                wtt.GetTable(SQLCode, startCol, endCol, startRow, destTabName);
          
                //System.Console.WriteLine("Error: The program failed when trying to write the data to Excel. Press any key to exit.");
                //System.Console.ReadLine();
                //Environment.Exit(1);
           
            try
            {
                //convert number col to num, use whatever column needs to be converted to a number format in excel
                RunSqlFromExcel.ConvertToNum ctn = new RunSqlFromExcel.ConvertToNum();
                ctn.ConvertMethod(numberRange, destTabName);
            }
            catch
            {
                System.Console.WriteLine("Error: The program failed when trying to convert a range to numbers in Excel. Press any key to exit.");
                System.Console.ReadLine();
                Environment.Exit(1);
            }

         

        }






    }
}
