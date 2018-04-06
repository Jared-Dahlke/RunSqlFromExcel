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
    class EndProc
    {



        public void EndProcMethod()
        {



            //Gets Excel and gets Activeworkbook and worksheet
            Excel.Application oXL;
            Excel.Workbook oWB;
            oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
            oXL.Visible = true;
            oWB = (Excel.Workbook)oXL.ActiveWorkbook;
            Excel.Sheets ExcelSheets = oWB.Worksheets;


            //refresh pivots
            // foreach (Microsoft.Office.Interop.Excel.PivotCache pt in oWB.PivotCaches())
            //   pt.Refresh();

            System.Console.WriteLine("Complete. Press enter to exit the program.");
            System.Console.ReadLine();

            return;
        }
        public string IsClientActive(string ClientName)
        {
            string SQLCode = "select ClientActive from ClientList where ClientName = '" + ClientName + "'";



            using (SqlConnection conn = new SqlConnection())
            {
                conn.ConnectionString = "Server = tcp:jadsolutions.database.windows.net,1433; Initial Catalog = Clients; Persist Security Info = False; User ID = jared.dahlke; Password = Qy1byeqy1bye; MultipleActiveResultSets = False; Encrypt = True; TrustServerCertificate = False; Connection Timeout = 30";

                try
                {
                    //initiate the connection
                    conn.Open();
                }
                catch (SqlException ex)
                {
                    throw new ApplicationException(string.Format("Check your internet connection and try again."), ex);
                }

                SqlCommand command = new SqlCommand(SQLCode, conn);

                // Create new SqlDataReader object and read data from the command.
                using (SqlDataReader reader = command.ExecuteReader())
                {
                    string clientStatus = null;
                    while (reader.Read())
                    {
                        clientStatus = reader["ClientActive"].ToString();
                    }

                    conn.Close();
                    return clientStatus;
                }
            }
        }



        class getConnStringClass
        {
            public string getEpicorConn()
            {
                string connString = null;
                connString = "Data Source=10.120.22.53;DATABASE=Epicor905;Workstation ID=SMEBPPL204TN;Trusted_Connection=true";
                return connString;
            }

            public string getTSWConn()
            {
                string connString = null;
                connString = "Data Source=10.120.22.52;DATABASE=TSWData;Workstation ID=SMEBPPL204TN;Trusted_Connection=true";
                return connString;
            }

            public string getAcctAutoFolder(string Company)
            {
                string connString = null;
                connString = "//10.120.22.15/WHG Corporate PCI/WHG Accounting/AccountingAutomation/FinancialTemplates - Do Not Modify/" + Company + ".xlsm";
                return connString;
            }

            public string getjellyFishSQLServerConn(string Company)
            {
                string connString = null;
                connString = "Server = tcp:jadsolutions.database.windows.net,1433; Initial Catalog = Clients; Persist Security Info = False; User ID = jared.dahlke; Password = Qy1byeqy1bye; MultipleActiveResultSets = False; Encrypt = True; TrustServerCertificate = False; Connection Timeout = 30";
                return connString;
            }

        }
    }
}
