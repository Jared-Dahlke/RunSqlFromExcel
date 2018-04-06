

						with CTE as (
						  SELECT 				      
					 CONVERT(VARCHAR(10), v.InvDate, 101) InvoiceDate,	
					 s.TeamName,
					 v.technician, 
					 v.TechnicianNumber,  
					 v.Cost
					 
					 
				   	From vw_JDAHLKE_SQL_DailyPartsUsage v  
					 left join mrc.serviceteamstechnicians t on v.TechnicianNumber = t.AgentNumber 
                          left join mrc.ServiceTeams s on t.STID = s.STID
						
                          
                          where (YEAR(v.CreateDate) = 2018 and Month(v.CreateDate) = 02 and (day(v.CreateDate) between 01 and 20)) )

						  SELECT
 *
FROM CTE
PIVOT (SUM(Cost) FOR [InvoiceDate] IN ([02/01/2018], [02/02/2018], [02/03/2018],[02/04/2018], [02/05/2018], [02/06/2018], [02/07/2018], [02/08/2018],[02/09/2018], [02/10/2018], [02/11/2018])) AS PIVOTED 



						--  group by 
						--  v.Technician,
						--  v.TechnicianNumber,
						--  s.teamName,
						--  v.invDate

						--  order by invoiceDate,
						--  s.teamName,
						--  v.Technician,
						--  v.TechnicianNumber,
					--	  cost




	

	