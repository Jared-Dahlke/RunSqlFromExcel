--declare @columnNames nvarchar(max) = ''
--select @columnNames += QUOTENAME(CONVERT(VARCHAR(10), InvDate, 101)) + ',' from vw_JDAHLKE_SQL_DailyPartsUsage where(YEAR(CreateDate) between   2018   and  2018   and Month(CreateDate) between  
--                           02   and   02   and (day(CreateDate) between   01   and   20  )) group by invdate
--set @columnNames = left(@columnNames, len(@columnNames)-1)


--DECLARE @query NVARCHAR(MAX)
--SET @query = '
--select * from
-- (SELECT 
-- CONVERT(VARCHAR(10), v.InvDate, 101) InvoiceDate, s.TeamName, v.technician, v.TechnicianNumber, v.Cost  
--                    From vw_JDAHLKE_SQL_DailyPartsUsage v  
--                    left join mrc.serviceteamstechnicians t on v.TechnicianNumber = t.AgentNumber  
--                          left join mrc.ServiceTeams s on t.STID = s.STID  
--                           where(YEAR(v.CreateDate) between   2018   and  2018   and Month(v.CreateDate) between  
--                           02   and   02   and (day(v.CreateDate) between   01   and   20  ))) as baseData
--                          PIVOT(SUM(Cost) FOR [InvoiceDate] IN( ' + @columnNames + '
--                         )) AS PIVOTED'

						 
--EXEC SP_EXECUTESQL @query



--select * from vw_JDAHLKE_SQL_ISForRevBySite where TransactionType like '%invoice%'

 --SELECT CONVERT(VARCHAR(10), date, 101) Date, BranchNumber, Period, TransactionType, Description, AccountNumber, Accountname, sum(Amount) as Amount, 
 --case when transactiontype in ('sales invoice', 'sales credit memo') then 'Chargeable Parts' 
 --     when transactiontype in ('service invoice', 'service credit memo') then 'Service Invoices' 
	--  when accountName in ('COGS-Service-Freight-In') then 'Freight' 
	--  when accountName in ('COGS-Service-PhysicalInvtyVariances') then 'Adjustments' 
 
 --else 'Other' end as ReportDef
 
 --From vw_JDAHLKE_SQL_ISForRevBySite 
 --where ((accountNumber like '%589100%' or accountNumber like '%589200%' or accountNumber like '%589400%'
 --                   or accountNumber like '%589500%' or accountNumber like '%588100%' or accountNumber like '%589600%' or accountNumber like '%581700%') AND 
 --                   ((Date >= '02/01/2018' and Date <= '02/20/2018') AND (BranchNumber >= '04' AND BranchNumber <= '95'))) 
 --                   Group by Date, BranchNumber, Period, TransactionType, Description, AccountNumber, Accountname, Amount


 --select Type, Date, BranchNumber, Amount
 --From( 
 --select 'Revenue' as Type, CONVERT(VARCHAR(10), date, 101) Date, BranchNumber, amount from vw_JDAHLKE_SQL_InternalSuppliesRev 
 --where (Date >= '02/01/2018' and Date <= '02/20/2018') and (AccountNumber like '%440099%' or AccountNumber like '%440199%' or AccountNumber like '%440299%'
 --or AccountNumber like '%440399%' or AccountNumber like '%441000%' or AccountNumber like '%441599%' or AccountNumber like '%441799%' or AccountNumber like '%445400%' or AccountNumber like '%446900%'
 --or AccountNumber like '%446900%' or AccountNumber like '%448000%' or AccountNumber like '%449700%' 
 --or (AccountNumber between '441800' and '442000') or (AccountNumber between '443000' and '443100') or (AccountNumber between '444000' and '444200') or (AccountNumber between '444600' and '444800')
 --or (AccountNumber between '445000' and '445100') or (AccountNumber between '449000' and '449100') or (AccountNumber between '443700' and '443800'))
 
 --UNION ALL

 --SELECT 'Expense' as Type, CONVERT(VARCHAR(10), Date, 101) Date, BranchNumber, Cost as amount
 --From vw_JDAHLKE_SQL_DailySuppliesExpense where orderType = 'SUPPLY' AND  
 --                   (Date >= '02/01/2018' and Date <= '02/20/2018') AND Amount = '0' AND InvoiceNumber > '1') x


				select * from SCBilling  --makes invoices

				select * from scbillingmeters  --meter type, equipmentID

				select* from scbillingmetergroups  ---invoice id, contractMeterGroup, 



					select * from scbillingequipments    --invoice id, equipID, 

					select * from arinvoices    --invoice id, invoiceNumber, bill to, customer id,



					SELECT ARC.CustomerNumber, ARC.CustomerName, ARI.InvoiceNumber, SCBMG.ContractMeterGroup, SCBM.MeterType AS MeterType 
		, SCBMG.CoveredCopies, SCBMG.CountedCopies, SCBMG.BillableCopies, SCBMG.EffectiveRate, SCBMG.AvgGroupRate, SCE.EquipmentNumber 
		, SCE.SerialNumber, SCBM.BeginMeterActual, SCBM.EndMeterActual, SCBM.DifferenceCopies FROM[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCBillingMeters SCBM 
        JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCBillingMeterGroups SCBMG ON SCBMG.BillingMeterGroupID = SCBM.BillingMeterGroupID 
       JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCBillingEquipments SCBE ON SCBE.EquipmentID = SCBM.EquipmentID AND SCBE.InvoiceID = SCBMG.InvoiceID 
        JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.ARInvoices ARI ON ARI.InvoiceID = SCBE.InvoiceID 
        JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.SCEQuipments SCE ON SCE.EquipmentID = SCBE.EquipmentID 
        JOIN[MRCEAUTO\SQLMRCEAUTO].coMRC.dbo.ARCustomers ARC ON ARC.CustomerID = ARI.CustomerID 
        WHERE ARI.InvoiceNumber LIKE 'IN810484' 
        ORDER BY SCBMG.ContractMeterGroup DESC



		select * from vw_JDAHLKE_SQL_AccountReviewVolumeBreakdown