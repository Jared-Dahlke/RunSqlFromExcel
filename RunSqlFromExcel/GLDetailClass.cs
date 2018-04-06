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
    class GLDetailClass
    {

        public void GLDetailMethod(string myReport, string myCol, string myLineItem, string myCompany, string myMonth, string myYear, string myPriorYear, string myReportName)
        {
            
           
            //find the activeWorkbook
            //Gets Excel and gets Activeworkbook and worksheet
            Excel.Application oXL;
            Excel.Workbook oWB;
            Excel.Worksheet excelWorksheet;
            oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Excel.Workbook)oXL.ActiveWorkbook;
            Excel.Sheets ExcelSheets = oWB.Worksheets;
            excelWorksheet = (Excel.Worksheet)ExcelSheets.get_Item("Raw Data");

           

            //get the last used row
            var lastUsedRow = excelWorksheet.Range["L2", "L900000"].get_End(XlDirection.xlDown).Row;


            //create an array of the excel input variables range                             
            Excel.Range variableRange = excelWorksheet.get_Range("D2", "K" + (lastUsedRow + 1));
            Object[,] VRO = (Object[,])variableRange.get_Value(Excel.XlRangeValueDataType.xlRangeValueDefault); //VRO is now a boxed two-dimensional array                                                                                                           

            
            // convert the object[] array to a Strongly Typed string array. Subtract 1 from VROs because the string array is 0 based, whereas the VRO  Object array is not zero based.
            string[,] VROs = new string[lastUsedRow, 10];
            for (int row = 1; row <= lastUsedRow; ++row)
            {
                for (int col = 1; col <= 8; ++col)
                {
                    if (VRO[row, col] != null)
                    {
                        VROs[row - 1, col - 1] = VRO[row, col].ToString();
                    }
                }
            }

            

            //establish a new list to put instances of my VROClass into:
            List<VROClass> VROClassList = new List<VROClass>();

            for (int row = 0; row <= lastUsedRow - 1; ++row)
            {
                if (VROs[row, 0] == myReportName && VROs[row, 1] == myLineItem)
                {
                    var rec = new VROClass
                    {
                        Report = VROs[row, 0],
                        lineItem = VROs[row, 1],
                        one = VROs[row, 2],
                        seg1 = VROs[row, 3],
                        seg2 = VROs[row, 4],
                        seg3 = VROs[row, 5],
                        seg4 = VROs[row, 6],
                        seg5 = VROs[row, 7]
                    };

                    VROClassList.Add(rec);

                }
               
            }



          //printToExcel(5000, VROs);



            //get the sql code to be run
            RunSqlFromExcel.getSQLClass sc = new RunSqlFromExcel.getSQLClass();
            string SQLCode = sc.getSQLGl(myCompany, myYear, myPriorYear, myMonth,myCol);



           

            System.Data.DataTable datatableSQL = new System.Data.DataTable();

            //write the sqlCode to dataTableSQL via adapter
            ConnStrings cs4 = new ConnStrings();
            string EpicorSQLConn = cs4.getConn("EpicorString", "0");

            using (SqlConnection conn = new SqlConnection(EpicorSQLConn))
            {
                SqlDataAdapter adapter = new SqlDataAdapter();
                adapter.SelectCommand = new SqlCommand(SQLCode, conn);
                conn.Open();
                adapter.Fill(datatableSQL);
                conn.Close();
            }

           
            //get a row and column count of datatablesql
            int recordsNum = datatableSQL.Rows.Count;
            int colNum = datatableSQL.Columns.Count;

            //convert the datatableSQL to a strongly typed list:
            string[,] SQLs = new string[recordsNum, colNum];
            for (int row = 0; row <= recordsNum - 1; ++row)
            {
                for (int col = 0; col <= colNum - 1; ++col)
                {

                    if (string.IsNullOrEmpty(datatableSQL.Rows[row][col].ToString()))
                    {
                        SQLs[row, col] = "0.00";
                    }
                    else
                    {

                        SQLs[row, col] = datatableSQL.Rows[row][col].ToString();
                    }
                }
            }

            
            // printToExcel(recordsNum, SQLs);
            
            var SQLClassList = new List<SQLClass>();
            //establish a new list to put instances of my SQLClass into:
            //List<SQLClass> SQLClassList = new List<SQLClass>();

            for (int row = 0; row <= recordsNum-1; ++row)
            {
                //load up a new record / instance of the SQLClass
                var record = new SQLClass
                {
                    segV1 = SQLs[row, 0].ToString(),
                    segV2 = SQLs[row, 1].ToString(),
                    segV3 = SQLs[row, 2].ToString(),                   
                    segV4 = SQLs[row, 3].ToString().ToUpper(),
                    segV5 = SQLs[row, 4].ToString().ToUpper(),

                    SQLFiscalYear = SQLs[row, 5].ToString(),
                    SQLFiscalPeriod = SQLs[row, 6].ToString(),
                    SQLDescription = SQLs[row, 7].ToString(),
                    SQLSourceModule = SQLs[row, 8].ToString(),
                    SQLJournalCode = SQLs[row, 9].ToString(),
                    SQLVendorNum = SQLs[row, 10].ToString(),
                    SQLAPInvoiceNum = SQLs[row, 11].ToString(),
                    SQLDebitAmount = SQLs[row, 12].ToString(),
                    SQLCreditAmount = SQLs[row, 13].ToString(),
                   
                };

                //add the instance of the SQLClass to the SQLClassList
                SQLClassList.Add(record);
            }
            

            //if seg 4 is null or empty then make it %
            foreach (var record in SQLClassList)
            {
                if (record.segV4 == "" || record.segV4 == null)
                {
                    record.segV4 = "%";

                }
                if (record.segV5 == "" || record.segV4 == null)
                {
                    record.segV5 = "%";

                }

            }




            // join 

            var final = from v in VROClassList
                        from s in SQLClassList
                        .Where(s =>
                            s.segV1.Like(v.seg1) &&
                            s.segV2.Like(v.seg2) &&
                            s.segV3.Like(v.seg3) &&
                            s.segV4.Like(v.seg4) &&
                            s.segV5.Like(v.seg5))//.DefaultIfEmpty()  // left join
                     //   group s by v into g
                        select new
                        {
                            segV1 = s.segV1,
                            segV2 = s.segV2,
                            segV3 = s.segV3,
                            segV4 = s.segV4,
                            segV5 = s.segV5,
                            FiscalYear = s.SQLFiscalYear,
                            FiscalPeriod = s.SQLFiscalPeriod,
                            Description = s.SQLDescription,
                            SourceModule = s.SQLSourceModule,
                            JournalCode = s.SQLJournalCode,
                            VendorNum = s.SQLVendorNum,
                            APInvoiceNum = s.SQLAPInvoiceNum,
                            DebitAmount = s.SQLDebitAmount,
                            CreditAmount = s.SQLCreditAmount

                            
                            
                        };

           

            string[,] finalTable = new string[1000000 + 1, 16];

           
            var i = 0;
            foreach (var row in final)
            {

                finalTable[i, 0] = row.segV1;
                finalTable[i, 1] = row.segV2;
                finalTable[i, 2] = row.segV3;
                finalTable[i, 3] = row.segV4;
                finalTable[i, 4] = row.segV5;
                finalTable[i, 5] = row.FiscalYear;
                finalTable[i, 6] = row.FiscalPeriod;
                finalTable[i, 7] = row.Description;
                finalTable[i, 8] = row.SourceModule;
                finalTable[i, 9] = row.JournalCode;
                finalTable[i, 10] = row.VendorNum;
                finalTable[i, 11] = row.APInvoiceNum;
                finalTable[i, 12] = row.DebitAmount;
                finalTable[i, 13] = row.CreditAmount;
                
                i++;
            }

            int iplus = i+1;


            //open new workbook
            //start new instance of excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            //open new workbook
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Sheets ExcelSheets2 = excelWorkbook.Worksheets;
            string currentSheet = "Sheet1";

            

            Excel.Worksheet excelWorksheet2 = (Excel.Worksheet)ExcelSheets2.get_Item(currentSheet);

            var tablePaste = excelWorksheet2.get_Range("a2", "n" + (recordsNum + 1));

            tablePaste.Value2 = finalTable;

            //input netAmount column
            var totalFormulaRange = excelWorksheet2.get_Range("o2", "o" + iplus).FormulaR1C1 = @"=+RC[-2]-RC[-1]";
           excelWorksheet2.get_Range("o2", "o" + iplus).Cells.NumberFormat = "#,##0.00";
            //input headers


            var headerRange = excelWorksheet2.get_Range("a1").Value2 = "Natural";
            var headerRange2 = excelWorksheet2.get_Range("b1").Value2 = "seg2";
            var headerRange3 = excelWorksheet2.get_Range("c1").Value2 = "seg3";
            var headerRange4 = excelWorksheet2.get_Range("d1").Value2 = "seg4";
            var headerRange5 = excelWorksheet2.get_Range("e1").Value2 = "seg5";
            var headerRange6 = excelWorksheet2.get_Range("f1").Value2 = "Year";
            var headerRange7 = excelWorksheet2.get_Range("g1").Value2 = "Month";
            var headerRange8 = excelWorksheet2.get_Range("h1").Value2 = "Description";
            var headerRange9 = excelWorksheet2.get_Range("i1").Value2 = "SourceModule";
            var headerRange10 = excelWorksheet2.get_Range("j1").Value2 = "Journal Code";
            var headerRange11 = excelWorksheet2.get_Range("k1").Value2 = "Vendor Number";
            var headerRange12 = excelWorksheet2.get_Range("l1").Value2 = "AP Invoice Num";
            var headerRange13 = excelWorksheet2.get_Range("m1").Value2 = "Debit Amount";
            headerRange13 = excelWorksheet2.get_Range("n1").Value2 = "Credit Amount";
            headerRange13 = excelWorksheet2.get_Range("o1").Value2 = "Net Amount";

            
            return;




        }



        private void printToExcel(int recordsNum, String[,] Results)
        {
            //start new instance of excel
            Excel.Application excelApp = new Excel.Application();
            excelApp.Visible = true;

            //open new workbook
            Excel.Workbook excelWorkbook = excelApp.Workbooks.Add();
            Excel.Sheets ExcelSheets = excelWorkbook.Worksheets;
            string currentSheet = "Sheet1";


            Excel.Worksheet excelWorksheet = (Excel.Worksheet)ExcelSheets.get_Item(currentSheet);

            var tablePaste = excelWorksheet.get_Range("a1", "m" + (recordsNum + 1));

            tablePaste.Value2 = Results;

            excelApp.DisplayFullScreen = true;


            Marshal.ReleaseComObject(ExcelSheets);


            return;

        }




    }




    public class VROClass
    {
        public string VROCompany { get; set; }
        public string Report { get; set; }
        public string lineItem { get; set; }
        public string one { get; set; }
        public string seg1 { get; set; }
        public string seg2 { get; set; }
        public string seg3 { get; set; }
        public string seg4 { get; set; }
        public string seg5 { get; set; }
    }


    public class SQLClass
    {
        public string SQLCompany { get; set; }
        public string segV1 { get; set; }
        public string segV2 { get; set; }
        public string segV3 { get; set; }
        public string segV4 { get; set; }
        public string segV5 { get; set; }
        public string SQLFiscalYear { get; set; }
        public string SQLFiscalPeriod { get; set; }
        public string SQLDescription{ get; set; }
        public string SQLSourceModule { get; set; }
        public string SQLJournalCode { get; set; }
        public string SQLVendorNum { get; set; }
        public string SQLAPInvoiceNum { get; set; }
        public string SQLDebitAmount { get; set; }
        public string SQLCreditAmount { get; set; }
       
    }

    class ConnStrings
    {
        public string getConn(string type, string Company)
        {
            string connString = null;
            if (type == "EpicorString")
            {
                connString = "Data Source=10.120.22.53;DATABASE=Epicor905;Workstation ID=SMEBPPL204TN;Trusted_Connection=true";
            }

            if (type == "TSWString")
            {
                connString = "Data Source=10.120.22.52;DATABASE=TSWData;Workstation ID=SMEBPPL204TN;Trusted_Connection=true";
            }

            if (type == "AccountingAutomationPLString")
            {
                connString = "//10.120.22.15/WHG Corporate PCI/WHG Accounting/AccountingAutomation/FinancialTemplates - Do Not Modify/" + Company + ".xlsm";
            }







            return connString;
        }


    }


    public static class QueryHelper
    {
        public static bool Like(this string target, string pattern)
        {
            if (target == null || pattern == null) return false;
            return WildcardToRegex(pattern).IsMatch(target);
        }

        public static Regex WildcardToRegex(string pattern)
        {
            return new Regex("^" + Regex.Escape(pattern)
            .Replace("%", ".*")
            .Replace("_", ".") + "$");
        }

    }

}
