<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Vinay Kumar
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>
<!--#include file="../includes/SiteConfig.asp"-->
<!--#include file="../includes/Connection_open.asp"-->
<%
	select case Request.Form("EditAction")
		case "MoveUp"
			for each approver in Request.Form("Purchaser")
				sql="sp_PSystem_SetPurchaserPriorityOneStepHigher '" & approver & "'"
				call DoSql(sql)
			next
		case "MoveDown"
			for each approver in Request.Form("Purchaser")
				sql="sp_PSystem_SetPurchaserPriorityOneStepLower '" & approver & "'"
				call DoSql(sql)
			next
		case "Delete"
			for each Purchaser in Request.Form("Purchaser")
				sql=sql_DeletePurchaser(Purchaser)
				call DoSql(sql)
			next
		case "Add"
			sql="sp_PSystem_AddPurchaser '" & Request.Form("Purchaser") & "'"
			call DoSQL(sql)
	end select
Response.Redirect("Purchasers.asp")
%>