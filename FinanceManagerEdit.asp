<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Subhash Rampuri
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>
<!--#include file="../includes/SiteConfig.asp"-->
<!--#include file="../includes/Connection_open.asp"-->
<%
	select case Request.Form("EditAction")
		case "Delete"
			for each FinanceManager in Request.Form("FinanceManager")
				sql=sql_DeleteFinanceManager(FinanceManager)
				call DoSql(sql)
			next
		case "Add"
			sql="sp_PSystem_AddFinanceManager '" & Request.Form("FinanceManager") & "'"
			call DoSQL(sql)
	end select
Response.Redirect("FinanceManager.asp")
%>