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
'----------Reject individual quotation code----------
'	Dim iloop,i
'	arrItems=split(Request.Form("ItemList"),"|")
'	for i = 0 to uBound(arrItems)
'		iCode = arrItems(i)
'		sql="select * from tbl_Psystem_Quotations where ItemCode = '" & iCode & "' "
'		call RunSql(sql,rsItems)
'
'		if not rsItems.EOF then
'			PrjId = rsItems("ProjectId")
'			ReqID = rsItems("RequisitionId")
'			IDesc = rsItems("ItemDescription")
'			ICode = rsItems("ItemCode")
'		sql = "Update tbl_PSystem_Quotations set isApproved = 3 where ProjectId = " & PrjId &" and RequisitionId = " & ReqID & " and ItemDescription = '" & Replace(Server.HTMLEncode(IDesc),"'","''") & "' and ItemCode = '" & ICode & "' "
'		Call DoSql(sql)
'	end if
'	next
'------------------------------------------------------
	ReqID = Request.form("hdReqNo")
	'Response.write ReqID

	sql = "Update tbl_Psystem_transactionDetails set Status = 3 where RequisitionId = " & ReqID & " and (Status = 1 or Status = 2)"
	Call DoSql(sql)

	sql = "Update tbl_Psystem_Quotations set isPRCancelled = 1 where RequisitionId = " & ReqID & "  and isApproved = 0 "
	Call DoSql(sql)

	%>

	<%
		Response.Redirect ("Purchase_QuotationList.asp")
	%>

<!--#include file="../includes/connection_close.asp"-->
