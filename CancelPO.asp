<%
'iMorfus Intranet Systems - Version 3.0.5 ' - Copyright 2002 - 04 (c) i-Vista Digital Solutions Limited.
'All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions.
'See the file iMorfuslicense.txt for more information.

'All Copyright notices must remain in place at all times.
'Developed By: Subhash Rampuri
'-----------------------------------------------------------------------------------------------

%>
<!--#include file="../includes/mail.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
 <%
	PurOrdNo = Request.form("hdPurOrdNo")
	Response.write PurOrdNo

	sql = "Update tbl_Psystem_Quotations set isPOCancelled = 1 where PurOrderNo = "& PurOrdNo &" "
	Call DoSql(sql)

	%>

	<%
		Response.Redirect ("PurchaseOrderReleased.asp")
	%>

<!--#include file="../includes/connection_close.asp"-->
