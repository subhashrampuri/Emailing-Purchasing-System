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
	ReqID = Request.form("hdReqNo")
	'Response.write ReqID

	sql = "Update tbl_Psystem_Quotations set isPROnHold = 1 where RequisitionId = " & ReqID & "  and isApproved = 0 "
	Call DoSql(sql)

	%>

	<%
		Response.Redirect ("Purchase_QuotationList.asp")
	%>

<!--#include file="../includes/connection_close.asp"-->
