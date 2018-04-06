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

namespace RunQueries.Console
{ 

    

    class Program
    {
       
        static void Main(string[] args)
        {
            RunSqlFromExcel.EndProc e = new RunSqlFromExcel.EndProc();
           if(e.IsClientActive("MRC") != "Active")
            {
                System.Console.WriteLine("Account requires maintenance. Please contact Jared to authenticate.");
                System.Console.ReadLine();
                Environment.Exit(1);
            }
           
            //collect args
            System.Console.WriteLine("Starting...");

            string beginPeriod = null;
            string endPeriod = null;
            string reportingPeriod = null;
            string report = null;

            try
            {
               beginPeriod = args.GetValue(0).ToString();
               endPeriod = args.GetValue(1).ToString();
               reportingPeriod = args.GetValue(2).ToString();
               report = args.GetValue(3).ToString();
            }
            catch
            {
                System.Console.WriteLine("Error: VBA variables passed in incorrectly. Press any key to exit.");
                System.Console.ReadLine();
                Environment.Exit(1);
            }
           

            if(report == "BillingAnalysis")
            {
                Billing(beginPeriod, endPeriod);
                return;
            }

            if (report == "SpecialInvoice")
            {
                SpecialInvoice(beginPeriod, report);
                return;
            }
            

            if (report == "BillToNext")
            {
                BillToNext(beginPeriod, report);
                return;
            }

            if (report == "SoftwareRenewalsReport")
            {
                NewSoftwareContracts(beginPeriod, endPeriod, report);
                return;
            }

            if (report == "ISForRevBySiteVereco")
            {
                ISForRevBySiteVereco(beginPeriod, endPeriod, report);
                return;
            }

            if (report == "BaseAudit")
            {
                BaseAudit(beginPeriod, endPeriod, report);
                return;
            }



            //sql name                                                       //tab name
            System.Console.WriteLine("Refreshing IS For Rev By Site report...");
            //DistributorMethod(query input 1, query input 2, name of SQL string that lives in the getSQLClass class, column where the query is placed, last column the query occupies,
            //starting row of query, range that needs to be converted to number in excel, siteID, reportingPeriod, name of the tab to be pasted on, report Name


            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, endPeriod, "ISForRevBySite", "A", "I", 2, "H:H","null", reportingPeriod, "ISServiceCOGS", "null");

            System.Console.WriteLine("Refreshing Supplies Expense Report...");
            dc.DistributorMethod(beginPeriod, endPeriod, "SuppliesExpense", "A", "N", 2, "L:L", "null", reportingPeriod, "SuppliesExpense", "null");

            System.Console.WriteLine("Refreshing Daily Parts Usage Report...");
           // dc.DistributorMethod(beginPeriod, endPeriod, "DailyPartsUsagePivot", "A", "AV", 2, "M:M", "null", reportingPeriod, "DailyPartsUsagePivot");
            dc.DistributorMethod(beginPeriod, endPeriod, "DailyPartsUsage", "A", "E", 2, "E:E", "null", reportingPeriod, "DailyPartsUsage", "null");

            System.Console.WriteLine("Refreshing Labor Cost Report...");
            dc.DistributorMethod(beginPeriod, endPeriod, "LaborUsage", "A", "E", 2, "E:E", "null", reportingPeriod, "LaborData", "null");

            System.Console.WriteLine("Refreshing Supplies Revenue Report...");
            dc.DistributorMethod(beginPeriod, endPeriod, "SuppliesRevenue", "A", "I", 2, "I:I", "null", reportingPeriod, "SuppliesRevenue", "null");

            System.Console.WriteLine("Refreshing Supplies Revenue and Expense Report...");
            dc.DistributorMethod(beginPeriod, endPeriod, "SuppliesRevAndExp", "A", "D", 2, "D:D", "null", reportingPeriod, "SuppliesRevAndExp", "null");

            System.Console.WriteLine("Refreshing Tech Linkage Report...");
            dc.DistributorMethod(beginPeriod, endPeriod, "TechLinkage", "A", "M", 2, "N:N", "null", reportingPeriod, "TechTeam", "null");

            System.Console.WriteLine("Refreshing Equipment Revenue by Branch Report...");
            dc.DistributorMethod(beginPeriod, endPeriod, "EquipRevBySite", "A", "C", 2, "C:C", "null", reportingPeriod, "EquipRevBySite", "null");

           


            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            
            


        }
        public static void Billing(string beginPeriod, string endPeriod)
        {
            string reportingPeriod = null;

            System.Console.WriteLine("Refreshing Billing Analysis report...");            
            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, endPeriod, "BillingAnalysis", "A", "X", 2, "Q:Q", "null", reportingPeriod, "BillingAnalysisData", "BillingAnalysisReport");


            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            return;
        }

        public static void SpecialInvoice(string beginPeriod, string report)
        {
            string reportingPeriod = null;

            System.Console.WriteLine("Compiling Invoice Data...");
            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, "null", "SpecialInvoice", "A", "Z", 2, "P:S", "null", reportingPeriod, "Invoice", report);  //begin period functioning as ivoice number


            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            return;
        }

        public static void BillToNext(string beginPeriod, string report)
        {
            string reportingPeriod = null;

            System.Console.WriteLine("Compiling SC Contracts Data...");
            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, "null", "SCContracts", "A", "D", 2, "E:E", "null", reportingPeriod, "Versants Data", report); //begin period functioning as month

            
            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            return;
        }

        public static void NewSoftwareContracts(string beginPeriod, string endPeriod, string report)
        {
            string reportingPeriod = null;

            System.Console.WriteLine("Compiling New Software Contracts Data...");
            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, endPeriod, "NewSoftwareContracts", "A", "M", 2, "E:E", "null", reportingPeriod, "New Software Contracts", report); //begin period functioning as month


            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            return;
        }

        public static void ISForRevBySiteVereco(string beginPeriod, string endPeriod, string report)
        {
            string reportingPeriod = null;

            System.Console.WriteLine("Compiling Vereco PL Data...");
            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, endPeriod, "ISForRevBySiteVereco", "A", "O", 2, "K:K", "null", reportingPeriod, "ISForRevBySiteVereco", report); //begin period functioning as month


            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            return;
        }

        public static void BaseAudit(string beginPeriod, string endPeriod, string report)
        {
            string reportingPeriod = null;

            System.Console.WriteLine("Compiling BaseAudit Data...");
            RunSqlFromExcel.DistributorClass dc = new RunSqlFromExcel.DistributorClass();
            dc.DistributorMethod(beginPeriod, endPeriod, "BaseAudit", "A", "B", 2, "K:K", "null", reportingPeriod, "BaseAudit", report); 


            RunSqlFromExcel.EndProc ep = new RunSqlFromExcel.EndProc();
            ep.EndProcMethod();
            return;
        }




    }
}
