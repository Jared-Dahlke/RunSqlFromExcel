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
    class getSQLClass
    {

        
        public string getSQLMethod(string begDateC, string endDateC, string codeType, string siteID, string reportingPeriod, string report)
        {
            //find the activeWorkbook
            //Gets Excel and gets Activeworkbook and worksheet

           
            Excel.Application oXL;
            Excel.Workbook oWB;
            try
            {

                oXL = (Excel.Application)Marshal.GetActiveObject("Excel.Application");
                oXL.Visible = true;
                oWB = (Excel.Workbook)oXL.ActiveWorkbook;
                Excel.Sheets ExcelSheets = oWB.Worksheets;
                Excel.Worksheet excelWorksheet = oWB.ActiveSheet;
            }
            catch
            {
                System.Console.WriteLine("Error: The program failed when trying to determine the active Excel sheet. Press any key to exit.");
                System.Console.ReadLine();
                Environment.Exit(1);

            }





            var begDate = begDateC;
            var endDate = endDateC;

            
            int begMonth = 0;
            int begDay = 0;
            int begYear = 0;

            int endMonth = 0;
            int endDay = 0;
            int endYear = 0;

            try
            {
                if (report != "BillingAnalysisReport" && report != "SpecialInvoice" && report != "BillToNext" && report != "ISForRevBySiteVereco"
                    && report != "BaseAudit" && report != "EquipRevBySite")
                {
                    begMonth = Convert.ToInt32(begDate.Split('/')[0]);
                    begDay = Convert.ToInt32(begDate.Split('/')[1]);
                    begYear = Convert.ToInt32(begDate.Split('/')[2]);
                    endMonth = Convert.ToInt32(endDate.Split('/')[0]);
                    endDay = Convert.ToInt32(endDate.Split('/')[1]);
                    endYear = Convert.ToInt32(endDate.Split('/')[2]);
                }
            }
            catch
            {

                System.Console.WriteLine("Error: The program failed when trying to parse parameters. Press any key to exit.");
                System.Console.ReadLine();
                Environment.Exit(1);

            }
           

          
           

            string SQLCode = null;
            if (codeType == "TSWAR")  
            {
                SQLCode = "select ar.SiteID, 'AR' as Type, ar.transactioncode, sum(ar.amount) as Amount FROM tswdata.dbo.t_arlineitem as ar WHERE((ar.posteddate <= '" + endDate + "' AND ar.posteddate >= '" + begDate + "') OR " +
                          "(ar.transactiondate <=  '" + endDate + "' AND ar.transactiondate >= '" + begDate + "' AND ar.posteddate IS NULL)) and ar.SiteID = " + siteID + " Group by ar.transactioncode, ar.siteid Order by ar.transactioncode, Amount, ar.siteid";
            }

            if (codeType == "ArFolioCodeGood") //TSWArFolio, this code runs AR and folio by transaction code and is used in the LWDO monthly tsw entry
            {
                SQLCode = "SELECT Co, lineType, tCode, Amt " +
                    "From(SELECT site.SiteID as Co, 'PortFolio' as lineType, FolioLineItem.TransactionCode as tCode, sum(FolioLineItem.Amount) as Amt " +
                    "FROM ((TSWDATA.dbo.Site Site INNER JOIN TSWDATA.dbo.FolioLineItem FolioLineItem ON Site.SiteID = FolioLineItem.SiteID) " +
                    "LEFT OUTER JOIN TSWDATA.dbo.Folio Folio ON FolioLineItem.FolioID = Folio.FolioID) " +
                    "LEFT OUTER JOIN TSWDATA.dbo.Reservation Reservation ON Folio.ReservationID = Reservation.ReservationID " +
                    "WHERE FolioLineItem.PostedDate IS NOT NULL AND FolioLineItem.PostedDate >= '" + begDate + "' AND FolioLineItem.PostedDate <= '" + endDate + "' AND Site.SiteID = " + siteID + " " +
                    "group by site.SiteID, foliolineitem.transactioncode " +
                    "UNION All " +
                    "Select ar.SiteID as Co, 'AR' as lineType, ar.transactioncode as tCode, sum(ar.amount) as Amt " +
                    "FROM tswdata.dbo.t_arlineitem as ar " +
                    "WHERE((ar.posteddate <= '" + endDate + "' AND ar.posteddate >= '" + begDate + "') OR (ar.transactiondate <= '" + endDate + "' AND ar.transactiondate >= '" + begDate + "' AND ar.posteddate IS NULL)) and ar.SiteID = "+siteID+ " "+
                    "Group by ar.transactioncode, ar.siteid) x Group By Co, lineType, tCode, Amt";
                
            }
            if (codeType == "ArFolioCode") //TSWArFolio, this code runs AR and folio by transaction code and is used in the LWDO monthly tsw entry
            {
                SQLCode = "SELECT Co, lineType, tCode, Amt " +
                    "From(SELECT site.SiteID as Co, 'PortFolio' as lineType, FolioLineItem.TransactionCode as tCode, sum(FolioLineItem.Amount) as Amt " +
                    "FROM ((TSWDATA.dbo.Site Site INNER JOIN TSWDATA.dbo.FolioLineItem FolioLineItem ON Site.SiteID = FolioLineItem.SiteID) " +
                    "LEFT OUTER JOIN TSWDATA.dbo.Folio Folio ON FolioLineItem.FolioID = Folio.FolioID) " +
                    "LEFT OUTER JOIN TSWDATA.dbo.Reservation Reservation ON Folio.ReservationID = Reservation.ReservationID " +
                    "WHERE FolioLineItem.PostedDate IS NOT NULL AND FolioLineItem.PostedDate >= '" + begDate + "' AND FolioLineItem.PostedDate <= '" + endDate + "' " +
                    "group by site.SiteID, foliolineitem.transactioncode " +
                    "UNION All " +
                    "Select ar.SiteID as Co, 'AR' as lineType, ar.transactioncode as tCode, sum(ar.amount) as Amt " +
                    "FROM tswdata.dbo.t_arlineitem as ar " +
                    "WHERE((ar.posteddate <= '" + endDate + "' AND ar.posteddate >= '" + begDate + "') OR (ar.transactiondate <= '" + endDate + "' AND ar.transactiondate >= '" + begDate + "' AND ar.posteddate IS NULL)) " +
                    "Group by ar.transactioncode, ar.siteid) x Group By Co, lineType, tCode, Amt";

            }

            if (codeType == "ISForRevBySite") // MRC
            {
                SQLCode = "SELECT CONVERT(VARCHAR(10), date, 101) Date, BranchNumber, Period, TransactionType, Description, " +
                    "AccountNumber, Accountname, sum(Amount) as Amount, "+
                    "case when transactiontype in ('sales invoice', 'sales credit memo') then 'Chargeable Parts' " +
                        "when transactiontype in ('service invoice', 'service credit memo') then 'Service Invoices' "+
                        "when accountName in ('COGS-Service-Freight-In') then 'Freight' "+
                        " when accountName in ('COGS-Service-PhysicalInvtyVariances') then 'Adjustments' "+
                        "else 'Other' end as ReportDef From vw_JDAHLKE_SQL_ISForRevBySite "+
                        "where ((accountNumber like '%589100%' or accountNumber like '%589200%' or accountNumber like '%589400%' " +
                        "or accountNumber like '%589500%' or accountNumber like '%588100%' or accountNumber like '%589600%' or accountNumber like '%581700%') AND " +
                        "((Date >= '" + begDateC + "' and Date <= '" + endDateC + "') AND (BranchNumber >= '04' AND BranchNumber <= '95'))) " +
                        "Group by Date, BranchNumber, Period, TransactionType, Description, AccountNumber, Accountname,  Amount";
            }

            if (codeType == "SuppliesExpense") // MRC
            {
                SQLCode = "SELECT SONumber, InvoiceNumber, InvoiceBillToName, InvoiceCustomerName, TransactionType, CONVERT(VARCHAR(10), Date, 101) Date, Period, " +
                    "ItemDescription, OrderType, Price, Amount, Cost, CreateDate, BranchNumber "+
                    "From vw_JDAHLKE_SQL_DailySuppliesExpense where orderType = 'SUPPLY' AND " +
                    "(Date >= '" + begDateC + "' and Date <= '" + endDateC + "') AND Amount = '0' AND InvoiceNumber > '1'";
            }

            if (codeType == "DailyPartsUsagePivot") // MRC
            {
                SQLCode = "with CTE as (SELECT CONVERT(VARCHAR(10), v.InvDate, 101) InvoiceDate, s.TeamName, v.technician, v.TechnicianNumber, v.Cost " +
                    "From vw_JDAHLKE_SQL_DailyPartsUsage v " +
                    "left join mrc.serviceteamstechnicians t on v.TechnicianNumber = t.AgentNumber " +
                          "left join mrc.ServiceTeams s on t.STID = s.STID " +
                           "where(YEAR(v.CreateDate) between " + begYear + " and " + endYear + " and Month(v.CreateDate) between " +
                           begMonth + " and " + endMonth + " and (day(v.CreateDate) between " + begDay + " and " + endDay + ")))" +
                          "SELECT * FROM CTE PIVOT(SUM(Cost) FOR[InvoiceDate] IN(" +
                          "[02 /01/2018], [02/02/2018], [02/03/2018],[02/04/2018], [02/05/2018], [02/06/2018], [02/07/2018], [02/08/2018],[02/09/2018], [02/10/2018], [02/11/2018])) AS PIVOTED";
            }

            if (codeType == "DailyPartsUsage") // MRC
            {
                SQLCode = "SELECT CONVERT(VARCHAR(10), v.InvDate, 101) InvoiceDate,	s.TeamName, v.technician, v.TechnicianNumber, sum(v.Cost) as Cost " +
                       "From vw_JDAHLKE_SQL_DailyPartsUsage v left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeamsTechnicians t on v.TechnicianNumber = t.AgentNumber " +
                          "left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeams s on t.STID = s.STID " +
                           "where(YEAR(v.CreateDate) between " + begYear + " and " + endYear + " and Month(v.CreateDate) between " +
                           begMonth + " and " + endMonth + " and (day(v.CreateDate) between " + begDay + " and " + endDay + "))" +
                         "group by v.Technician, v.TechnicianNumber, s.teamName, v.invDate " +
                         "order by invoiceDate, s.teamName, v.Technician, v.TechnicianNumber, Cost";

            }

            if (codeType == "TechLinkage") // MRC
            {
                SQLCode = "SELECT * from [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeamsTechnicians t " +
                          "left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeams s on t.STID = s.STID ";

            }
           // SELECT* FROM[MRCSQL1].Digital_hub.dbo.ACC_ServiceTeams
      // SELECT* FROM[MRCSQL1].Digital_hub.dbo.ACC_ServiceTeamsTechnicians

            if (codeType == "LaborUsage") // MRC
            {
                SQLCode = "SELECT CONVERT(VARCHAR(10), v.InvDate, 101) InvoiceDate,	s.TeamName, v.technician, v.TechnicianNumber, sum(v.LaborCost) as LaborCost " +
                       "From vw_JDAHLKE_SQL_SCWSNCustNameTravelLabor v " +
                       "left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeamsTechnicians t on v.TechnicianNumber = t.AgentNumber " +
                         "left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeams s on t.STID = s.STID " +

                           "where(YEAR(InvDate) between " + begYear + " and " + endYear + " and Month(InvDate) between " +
                           begMonth + " and " + endMonth + " and (day(InvDate) between " + begDay + " and " + endDay + "))" +
                           "group by v.Technician, v.TechnicianNumber, s.teamName, v.invDate " +
                         "order by invoiceDate, s.teamName, v.Technician, v.TechnicianNumber, LaborCost";

            }

            if (codeType == "DailyPartsUsageAll") // MRC
            {
                SQLCode = "SELECT *  " +
                       "From vw_JDAHLKE_SQL_DailyPartsUsage v left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeamsTechnicians t on v.TechnicianNumber = t.AgentNumber " +
                          "left join [MRCSQL1].Digital_hub.dbo.ACC_ServiceTeams s on t.STID = s.STID " +
                           "where(YEAR(v.CreateDate) between " + begYear + " and " + endYear + " and Month(v.CreateDate) between " +
                           begMonth + " and " + endMonth + " and (day(v.CreateDate) between " + begDay + " and " + endDay + "))";
                         

            }

            if (codeType == "SuppliesRevenue") // MRC
            {
                SQLCode = "select TransID, BranchNumber, CONVERT(VARCHAR(10), date, 101) Date, Period, " +
                    "TransactionType, Description, AccountNumber, AccountName,  amount from vw_JDAHLKE_SQL_InternalSuppliesRev " +
                    "where(Date >= '" + begDate + "' and Date <= '" + endDate + "') and " +
                    "(AccountNumber like '%440099%' or AccountNumber like '%440199%' or AccountNumber like '%440299%' " +
                    "or AccountNumber like '%440399%' or AccountNumber like '%441000%' or AccountNumber like '%441599%' or AccountNumber like " +
                    "'%441799%' or AccountNumber like '%445400%' or AccountNumber like '%446900%' or AccountNumber like '%446900%' or " +
                    "AccountNumber like '%448000%' or AccountNumber like '%449700%' or(AccountNumber between '441800' and '442000') " +
                    "or(AccountNumber between '443000' and '443100') or(AccountNumber between '444000' and '444200') or " +
                    "(AccountNumber between '444600' and '444800') or(AccountNumber between '445000' and '445100') or " +
                    "(AccountNumber between '449000' and '449100') or(AccountNumber between '443700' and '443800'))";

                
            }
            if (codeType == "SuppliesRevAndExp") // MRC
            {
                SQLCode = "select Type, Date, BranchNumber, Amount " +
                          "From( select 'Revenue' as Type, CONVERT(VARCHAR(10), date, 101) Date, BranchNumber, amount " +
                          "from vw_JDAHLKE_SQL_InternalSuppliesRev " +
                            "where (Date between '" + begDate + "' and '" + endDate + "') and(AccountNumber like '%440099%' or AccountNumber like '%440199%' or " +
                            "AccountNumber like '%440299%' or AccountNumber like '%440399%' or AccountNumber like '%441000%' or " +
                            "AccountNumber like '%441599%' or AccountNumber like '%441799%' or AccountNumber like '%445400%' or AccountNumber like '%446900%' " +
                            "or AccountNumber like '%446900%' or AccountNumber like '%448000%' or AccountNumber like '%449700%' or(AccountNumber " +
                            "between '441800' and '442000') or(AccountNumber between '443000' and '443100') or(AccountNumber between '444000' and '444200') " +
                            "or(AccountNumber between '444600' and '444800') or(AccountNumber between '445000' and '445100') or " +
                            "(AccountNumber between '449000' and '449100') or(AccountNumber between '443700' and '443800')) " +
                            "UNION ALL SELECT 'Expense' as Type, CONVERT(VARCHAR(10), Date, 101) Date, BranchNumber, Cost as amount " +
                            "From vw_JDAHLKE_SQL_DailySuppliesExpense where orderType = 'SUPPLY' AND " +
                            "(Date between '" + begDate + "' and '" + endDate + "') AND Amount = '0' AND InvoiceNumber > '1') x";
            }


            if (codeType == "BillingAnalysis") // MRC
            {
                SQLCode = "select * from vw_JDAHLKE_SQL_BillingAnalysisByContract where invoiceperiod between '"+begDate+"' and '"+endDate+"'";

            }

            if (codeType == "Invoice") // MRC  //legacy invoice sql , developed before i found out they use views to pull special invoices
            {
                SQLCode = "SELECT ARC.CustomerNumber, ARC.CustomerName, ARI.InvoiceNumber, SCBMG.ContractMeterGroup, SCBM.MeterType AS MeterType "+
		", SCBMG.CoveredCopies, SCBMG.CountedCopies, SCBMG.BillableCopies, SCBMG.EffectiveRate, SCBMG.TotalChargeAmount, SCE.EquipmentNumber "+
		@", SCE.SerialNumber, SCBM.BeginMeterActual, SCBM.EndMeterActual, SCBM.DifferenceCopies FROM [MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCBillingMeters SCBM "+

        @"JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCBillingMeterGroups SCBMG ON SCBMG.BillingMeterGroupID = SCBM.BillingMeterGroupID "+

       @"JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCBillingEquipments SCBE ON SCBE.EquipmentID = SCBM.EquipmentID AND SCBE.InvoiceID = SCBMG.InvoiceID "+

        @"JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.ARInvoices ARI ON ARI.InvoiceID = SCBE.InvoiceID "+

        @"JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCEQuipments SCE ON SCE.EquipmentID = SCBE.EquipmentID "+

        @"JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.ARCustomers ARC ON ARC.CustomerID = ARI.CustomerID "+

        "WHERE ARI.InvoiceNumber LIKE '"+ begDateC + "' "+

        "ORDER BY SCBMG.ContractMeterGroup DESC";

            }

            if (codeType == "SpecialInvoice") //MRC
            {
                SQLCode = "SELECT * from vw_JDAHLKE_SQL_AccountReviewVolumeBreakdown WHERE InvoiceNumber LIKE '" + begDateC + "'";

            }

            if (codeType == "SCContracts") //MRC
            {
                //SQLCode = "select contractNumber, CONVERT(VARCHAR(10), BaseNextBillingDate, 101) date from SCContracts " +
                //          "where  Month(BaseNextBillingDate) = '" + begDate + "'";


                SQLCode = "select s.contractNumber, CONVERT(VARCHAR(10), s.BaseNextBillingDate, 101) date, v.Model, v.serialNumber from vw_JDAHLKE_SQL_ServiceContractsEquip " +
          "v left join scContracts s on s.contractNumber = v.ContractNumber " +
          "where month(s.BaseNextBillingDate) = '" + begDate + "' and(v.Model like '%V80B' or v.Model like '%V180B%' or v.Model like '%V180P%' or v.Model like '%V2100%' or v.Model like '%dp 180RC%' or v.Model like '%XV80P%') "+
          "group by  s.BaseNextBillingDate, s.contractNumber, v.Model, v.serialNumber" ;


            }
            if (codeType == "NewSoftwareContracts") //MRC  //not in use
            {
                //SQLCode = "select contractNumber, CONVERT(VARCHAR(10), BaseNextBillingDate, 101) date from SCContracts " +
                //          "where  Month(BaseNextBillingDate) = '" + begDate + "'";


                SQLCode = "select ContractNumber, CustomerName, CustomerNumber, StartDate, ExpDate, BaseRate, ponumber, ActivatedDate " +
                    "from vw_JDAHLKE_SQL_ServiceContractsClear where contractCode like '%software%' and CONVERT(VARCHAR(10), Activated, 101) ActivatedDate " +
                    "between '12/13/2017' and '03/16/2018'";


            }


            if (codeType == "ISForRevBySiteVereco") // MRC
            {
                SQLCode = "select v.Reference, v.TransactionType, v.CreatorID, v.AccountNumber, " +
                    "v.CustomerName as ISCustomerName, v.Description, v.Description2, v.BranchNumber, " +
        "i.CustomerID, c.CustomerName as ARCustCustomerName, " +
       "v.Amount, lineitemlink.LineItem, l.PLCustomerName as LinkageCustomerName, " +
        "case "+
        //"When v.Description like '%LOMA%' then 'Loma Linda' " +
//        " When v.Description like '%LLU%' then 'Loma Linda' " +
//"When v.Description like '%MOUNTAIN VIEW%' then 'Loma Linda' " +
//"When v.Description like '%Thomas Ziska%' then 'Loma Linda' " +
//"When v.Description like '%Chris Ming%' then 'Loma Linda' " +
//"When v.Description like '%24747 REDLANDS%' then 'Loma Linda' " +
//"When v.Description like '%Arbors Business Center%' then 'Loma Linda' " +
//"When v.Description like '%City of Hope%' then 'COH' " +
//"When v.Description like '%CityofHope%' then 'COH' " +
//"When v.Description like '%COH%' then 'COH' " +
//"When v.Description like '%UHS %' then 'UHS' " +
//"When v.Description like '%TCMC%' then 'UHS' " +
//"When v.Description like '%STHS%' then 'UHS' " +
//"When v.Description like '%TEXOMA%' then 'UHS' " +
//"When v.Description like '%South Texas Health%' then 'UHS' " +
//"When v.Description like '%Southwest Health%' then 'SW' " +
//"When v.Description like '%St Marys Region%' then 'UHS' " +
//"When v.Description like '%BRIAN WILLIAMS%' then 'UHS' " +
//"When v.Description like '%ARTHUR TRUJILLO%' then 'UHS' " +
//"When v.Description like '%Wellington Regional%' then 'UHS' " +
//"When v.Description like '%SWHS%' then 'UHS' " +
"When v.Description like '%EDGARRENDON%' then 'UHS' " +
"When v.Description like '%alejandroCardenas%' then 'LomaLinda' " +
//"When v.Description like '%Riverside Univ%' then 'Riverside' " +
//"When v.Description like '%RUHS%' then 'Riverside' " +
//"When v.Description like '%TRI CITY%' then 'TRI CITY' " +
//"When v.Description like '%HMH%' then 'HMH' " +
//"When v.Description like '%Palmdale Regional%' then 'Palmdale' " +
////"When v.Description like '%Vereco%' then 'Loma Linda' " +

"When v.Description2 like '%LOMA%' then 'Loma Linda' " +
"When v.Description2 like '%LLU%' then 'Loma Linda' " +
"When v.Description2 like '%MOUNTAIN VIEW%' then 'Loma Linda' " +
"When v.Description2 like '%Thomas Ziska%' then 'Loma Linda' " +
"When v.Description2 like '%Chris Ming%' then 'Loma Linda' " +
"When v.Description2 like '%24747 REDLANDS%' then 'Loma Linda' " +
"When v.Description2 like '%City of Hope%' then 'COH' " +
"When v.Description2 like '%CityofHope%' then 'COH' " +
"When v.Description2 like '%COH%' then 'COH' " +
"When v.Description2 like '%UHS %' then 'UHS' " +
"When v.Description2 like '%TCMC%' then 'UHS' " +
"When v.Description2 like '%STHS%' then 'UHS' " +
"When v.Description2 like '%TEXOMA%' then 'UHS' " +
"When v.Description2 like '%South Texas Health%' then 'UHS' " +
"When v.Description2 like '%Southwest Health%' then 'SW' " +
"When v.Description2 like '%St Marys Region%' then 'UHS' " +
"When v.Description2 like '%BRIAN WILLIAMS%' then 'UHS' " +
"When v.Description2 like '%ARTHUR TRUJILLO%' then 'UHS' " +
"When v.Description2 like '%SWHS%' then 'UHS' " +
"When v.Description2 like '%EDGAR RENDON%' then 'UHS' " +
"When v.Description2 like '%Riverside Univ%' then 'Riverside' " +
"When v.Description2 like '%Riverside%' then 'Riverside' " +
"When v.Description2 like '%RUHS%' then 'Riverside' " +
"When v.Description2 like '%TRI CITY%' then 'TRI CITY' " +
"When v.Description2 like '%HMH%' then 'HMH' " +
//"When v.Description2 like '%Vereco%' then 'Loma Linda' " +
"When c.CustomerName like '%LOMA%' and c.Customername NOT LIKE '%Murr%' then 'Loma Linda' " +
"When c.CustomerName like '%Murrieta%' then 'MURRIETA' " +
//newlyadded
"When v.CustomerName like '%LOMA%' and c.Customername NOT LIKE '%Murr%' then 'Loma Linda' " +
"When v.CustomerName like '%Murrieta%' then 'MURRIETA' " +

"When c.CustomerName like '%South Texas%' then 'UHS' " +
"When v.customerName like '%Palmdale Regional%' then 'Palmdale' " +
"When v.reference like '%DM01937%' then 'LomaLinda' " +


"else l.PLCustomerName " +
"end as 'PLCustomerNameFinal', v.period " +

"from vw_JDAHLKE_SQL_ISForRevBySite v " +

"left join ARInvoices i on v.Reference = i.InvoiceNumber " +
"left join ARCustomers c on i.CustomerID = c.CustomerID " +
"left join (select * from [MRC].['NAME FOR PL$'] group by EautoCustomerName, PLCustomerName) l on c.CustomerName = l.EautoCustomerName " +
"left join[MRC].['GL LOOKUP (2)$'] lineItemLink on v.AccountNumber = lineItemLink.GLAccount " +
//"left join (select voucherNumber, voucherID from APVouchers group by voucherNumber, voucherID) a on v.Reference = a.voucherNumber "+
//"left join (select voucherID, description, amount from APVoucherDetails group by voucherID, description, amount) ad on a.voucherid = ad.VoucherID "+
"where (v.Period between '" + begDate + "' and '" + endDate + "') and v.AccountNumber > 399999 and V.accountNumber <> 560002 " +
"and v.BranchNumber = '05' " +

"UNION ALL " +

"select v.Reference,  v.TransactionType, v.CreatorID, v.AccountNumber, v.CustomerName as ISCustomerName, " +
"v.Description, v.Description2, v.BranchNumber,  " +
 " i.CustomerID, c.CustomerName as ARCustCustomerName, v.Amount, lineitemlink.LineItem, l.PLCustomerName as LinkageCustomerName, " +

"l.PLCustomerName as 'PLCustomerNameFinal', v.period " +

"from vw_JDAHLKE_SQL_ISForRevBySite v " +
"left join ARInvoices i on v.Reference = i.InvoiceNumber " +
"left join ARCustomers c on i.CustomerID = c.CustomerID " +
"left join (select * from [MRC].['NAME FOR PL$']  group by EautoCustomerName, PLCustomerName) l on c.CustomerName = l.EautoCustomerName " +
"left join[MRC].['GL LOOKUP (2)$'] lineItemLink on v.AccountNumber = lineItemLink.GLAccount " +
//"left join (select voucherNumber, voucherID from APVouchers group by voucherNumber, voucherID) a on v.Reference = a.voucherNumber " +
//"left join (select voucherID, description, amount from APVoucherDetails group by voucherID, description, amount) ad on a.voucherid = ad.VoucherID " +
"where  v.Period between '" + begDate + "' and '" + endDate + "' and v.AccountNumber > 399999 and V.accountNumber <> 560002 and v.BranchNumber <> '05'";



            }

            if (codeType == "BaseAudit") //MRC 
            {
               

                SQLCode = "select s.ContractNumber, s.CustomerName from vw_JDAHLKE_SQL_ServiceContracts s " +
                            "left join(select coveredcopies, ContractNumber from vw_JDAHLKE_SQL_ServiceContractBillingMeterGroups group by CoveredCopies, ContractNumber) m " +
                            "on s.ContractNumber = m.ContractNumber " +
                            "where s.BaseRate = 0 and m.CoveredCopies > 0 and s.active = 1 " +
                            "group by s.ContractNumber, s.CustomerName";

                
            }

            if (codeType == "EquipRevBySite") //MRC 
            {

                SQLCode = "select v.BranchNumber, v.Period, sum(v.Amount) from vw_JDAHLKE_SQL_ISForRevBySite v " +

                      "where(v.AccountNumber in (400100, 401000, 401199, 402000, 403000, 404000, 407000, 408000, 409700, 409500) "+
                      "or v.AccountNumber between 400000 and 400002 " +
                      "or v.AccountNumber between 400300 and 400800 " +
                      "or v.AccountNumber between 401200 and 401300 " +
                      "or v.AccountNumber between 401500 and 401600 " +
                      "or v.AccountNumber between 402500 and 402700 " +
                      "or v.AccountNumber between 405000 and 406600 " +
                      "or v.AccountNumber between 409000 and 409300) and v.Period like '"+reportingPeriod+"' " +
                      "group by v.BranchNumber, v.Period";

            }



            









            return SQLCode;
        }




        public string getSQLGl(string CompanyVar, string FiscalYear, string FiscalYearPYY, string FiscalPeriod, string myCol)
        {
            string SQLCode = null;
            if(myCol == "15") //actual YTD
            {

               SQLCode = "SELECT GLJrnDtl.segValue1, GLJrnDtl.segValue2, GLJrnDtl.segValue3, GLJrnDtl.segValue4, GLJrnDtl.segValue5, GLJrnDtl.FiscalYear, GLJrnDtl.FiscalPeriod, " +
                    "GLJrnDtl.Description, GLJrnDtl.SourceModule, GLJrnDtl.JournalCode, GLJrnDtl.VendorNum, GLJrnDtl.APInvoiceNum, GLJrnDtl.DebitAmount, " +
                    "GLJrnDtl.CreditAmount " +
                    "FROM Epicor905.dbo.GLJrnDtl GLJrnDtl " +
                    "WHERE GLJrnDtl.Company = '" + CompanyVar + "' and GLJrnDtl.FiscalYear = '" + FiscalYear + "' and GLJrnDtl.FiscalPeriod <= '" + FiscalPeriod + "'";
               

            }
            if (myCol == "23") //prior year ytd
            {

                SQLCode = "SELECT GLJrnDtl.segValue1, GLJrnDtl.segValue2, GLJrnDtl.segValue3, GLJrnDtl.segValue4, GLJrnDtl.segValue5, GLJrnDtl.FiscalYear, GLJrnDtl.FiscalPeriod, " +
                     "GLJrnDtl.Description, GLJrnDtl.SourceModule, GLJrnDtl.JournalCode, GLJrnDtl.VendorNum, GLJrnDtl.APInvoiceNum, GLJrnDtl.DebitAmount, " +
                     "GLJrnDtl.CreditAmount " +
                     "FROM Epicor905.dbo.GLJrnDtl GLJrnDtl " +
                     "WHERE GLJrnDtl.Company = '" + CompanyVar + "' and GLJrnDtl.FiscalYear = '" + FiscalYearPYY + "' and GLJrnDtl.FiscalPeriod <= '" + FiscalPeriod + "'";


            }
            if (myCol == "2") //cm
            {

                SQLCode = "SELECT GLJrnDtl.segValue1, GLJrnDtl.segValue2, GLJrnDtl.segValue3, GLJrnDtl.segValue4, GLJrnDtl.segValue5, GLJrnDtl.FiscalYear, GLJrnDtl.FiscalPeriod, " +
                     "GLJrnDtl.Description, GLJrnDtl.SourceModule, GLJrnDtl.JournalCode, GLJrnDtl.VendorNum, GLJrnDtl.APInvoiceNum, GLJrnDtl.DebitAmount, " +
                     "GLJrnDtl.CreditAmount " +
                     "FROM Epicor905.dbo.GLJrnDtl GLJrnDtl " +
                     "WHERE GLJrnDtl.Company = '" + CompanyVar + "' and GLJrnDtl.FiscalYear = '" + FiscalYear + "' and GLJrnDtl.FiscalPeriod = '" + FiscalPeriod + "'";

            }
            if (myCol == "10") //pym
            {

                SQLCode = "SELECT GLJrnDtl.segValue1, GLJrnDtl.segValue2, GLJrnDtl.segValue3, GLJrnDtl.segValue4, GLJrnDtl.segValue5, GLJrnDtl.FiscalYear, GLJrnDtl.FiscalPeriod, " +
                     "GLJrnDtl.Description, GLJrnDtl.SourceModule, GLJrnDtl.JournalCode, GLJrnDtl.VendorNum, GLJrnDtl.APInvoiceNum, GLJrnDtl.DebitAmount, " +
                     "GLJrnDtl.CreditAmount " +
                     "FROM Epicor905.dbo.GLJrnDtl GLJrnDtl " +
                     "WHERE GLJrnDtl.Company = '" + CompanyVar + "' and GLJrnDtl.FiscalYear = '" + FiscalYearPYY + "' and GLJrnDtl.FiscalPeriod = '" + FiscalPeriod + "'";

            }



            return SQLCode;
        }

       

    }
}
