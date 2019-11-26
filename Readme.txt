
Hi,
After updating Tables and stored procedures,
Run below Sql Update statements 

/*
	Update tbl_Psystem_PurchaseRequestMaster set RequisitionNum = RequisitionId
	
	Update tbl_Psystem_PurchaseOrder set PurOrderNum = PurOrderNo	
	
	Update tbl_Psystem_GRN set GRNNum = GRNNo	

*/

--job schedular stored Procedure---

CREATE PROCEDURE [dbo].[sp_itbl_PSystem_Control] 
as
	insert into tbl_Psystem_Control values ('PR',dbo.fn_PSystem_GetFinancialYear(getdate()),Cast(Year(GETDATE()) +1 as varchar(10))+'-04-01',0)
	insert into tbl_Psystem_Control values ('PO',dbo.fn_PSystem_GetFinancialYear(getdate()),Cast(Year(GETDATE()) +1 as varchar(10))+'-04-01',0)
	insert into tbl_Psystem_Control values ('GRN',dbo.fn_PSystem_GetFinancialYear(getdate()),Cast(Year(GETDATE())+1 as varchar(10))+'-04-01',0)
GO
