<%@ LANGUAGE="VBSCRIPT" %>
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
	Dim iloop,i
 	'Response.write Request.Form("ItemList")
	arrItems=split(Request.Form("ItemList"),"|")

	for i = 0 to uBound(arrItems)
		iCode = arrItems(i)
		sql="select * from tbl_Psystem_Quotations where ItemCode = '" & iCode & "' "
		call RunSql(sql,rsItems)

		if not rsItems.EOF then
			'Response.write (rsItems("RequisitionId")) & "<br>"
			'Response.write (rsItems("ProjectId")) & "<br>"
			'Response.write (rsItems("ItemDescription")) & "<br>"
			'Response.write (rsItems("ItemCode")) & "<br>"

			PrjId = rsItems("ProjectId")
			ReqID = rsItems("RequisitionId")
			IDesc = rsItems("ItemDescription")
			ICode = rsItems("ItemCode")

		sql = "Update tbl_PSystem_Quotations set isApproved = 1,isPROnHold = 0 where ProjectId = " & PrjId &" and RequisitionId = " & ReqID & " and ItemDescription = '" & Replace(Server.HTMLEncode(IDesc),"'","''") & "' and ItemCode = '" & ICode & "' "
		'Response.write sql
		Call DoSql(sql)

		sql = "Update tbl_PSystem_Quotations Set isApproved = 2,isPROnHold = 0 where RequisitionId = " & ReqId & " and ProjectId = " & PrjID & " and ItemDescription = '" & Replace(Server.HTMLEncode(IDesc),"'","''") & "' and isApproved = 0 "
		Call DoSql(sql)

			'sql = "Update tbl_PSystem_PurchaseRequestTransaction Set Status = 5 where ProjectId = " & PrjID & " and RequisitionId = " & ReqID & " and ItemDescription = '" & Replace(Server.HTMLEncode(IDesc),"'","''") & "' "
			'Call DoSql(sql)

		sql = "Update tbl_Psystem_TransactionDetails set Status = 4 where ProjectId = " & PrjID & " and RequisitionId = " & ReqID & " and ItemDescription = '" & Replace(Server.HTMLEncode(IDesc),"'","''") & "' "
		Call DoSql(sql)


		end if
	Next

	sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& ReqID &" "
	call RunSql(sql,rsRec)
	if rsRec.Eof = false then
		ReqNum = rsRec("RequisitionNum")
	end if
	rsRec.Close

	Call fsSendMail_PurchaseTeam()
	%>
	<%
	Private Function fsSendMail_PurchaseTeam()
	Dim sBody
	'-----Active Approver---------------
	sql="sp_PSystem_GetActiveApprover"
	call RunSql(sql,rsApprover)
	if rsApprover.eof then
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in Approvers Panel.</b></font></center>"
		Response.end
	end if
	ApproverId=rsApprover("EmployeeId")
	eFromName=rsApprover("EmployeeName")
	eFromEmail=rsApprover("EmployeeEmail")
	rsApprover.close()
	'-----Active Purhase Team------------
	sql = " select EmployeeID from tbl_PSystem_PurchaseTeam "
	call RunSql(sql,rsEmp)

	if rsEmp.Eof = false then
		While NOT rsEmp.EOF
		EmpID = rsEmp("EmployeeID")

	str = "select dbo.fn_TSystem_EmployeeName('" & EmpID & "') as Name, dbo.fn_TSystem_EmployeeEmail('" & EmpID & "') as Email"
  	Call RunSql(str,rsPur)

	eToName =  rsPur("Name")
	eToEmail =  rsPur("Email")

	sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
			"</b><br>" & _
			" The synopses of the approved best quotations are as follows: " & _
			"</font><br><br>" & _
	 		"<table width='100%' border='0' cellspacing='2' cellpadding='2'>"  & _
		 	" <tr bgcolor=#108ed6>  " & _
			" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></div> </td>" & _
		  	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>ItemDescription</b></font> </div> </td> " & _
		 	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Supplier Name</b></font> </div> </td> " & _
		 	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Unit Price</b></font> </div> </td>" & _
		  	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Tax</b></font> </div> </td>" & _
		  	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Tax Percent</b></font> </div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity</b></font> </div> </td> " & _
		  	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Warranty</b></font> </div> </td>" & _
		    " <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Delivery Time</b></font></div></td>" & _
			" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Payment Terms</b></font></div> </td>" & _
		 	" </tr> "
		 if ReqID <> "" then
		  sql = " Select * from tbl_Psystem_Quotations where isApproved = 1 and RequisitionId = "& ReqID &" "
		  Call RunSql(sql,rsReq)
		  i = 1
		  while not rsReq.EOF
  			if rsReq("Currency") = -1 then
				sCurr = "Rupee(s)"
			else
				sCurr = "Doller(s)"
			end if
			if rsReq("isTaxIncludedOrExcluded") = -1 then
				sTax = "Exclusive"
			else
				sTax = "Inclusive"
			end if


	sBody  = sBody &  "<tr bgcolor=#DFF2FC>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& i &"</font></div> </td>" & _
			" <td> <div align='left'><font face='Trebuchet MS'>"& rsReq("ItemDescription") &"</font></div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& rsReq("SupplierName") &"</font></div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& rsReq("UnitPrice") & " " & sCurr &"</font></div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& sTax &"</font></div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>" & rsReq("TaxPercent") & " " & "%" & "</font></div> </td> " & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& rsReq("Quantity")&"</font> </div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& rsReq("Warranty") &"</font> </div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& rsReq("DeliveryTime") &"</font> </div> </td>" & _
			" <td> <div align='center'><font face='Trebuchet MS'>"& rsReq("PaymentTerms")&"</font> </div> </td>" & _
		    " </tr> "
			i = i + 1
			rsReq.movenext
			Wend
			rsReq.Close
		  end if
	sBody = sBody &	"<tr bgcolor=#108ed6>" &_
			"<td colspan='10' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail.</font></td>" &_
			"</tr>" &_
		    "</table>"

		'Response.write sBody
		eSubject = "Approved Purchase Quotation  : " & GetPurchaseRequisitionNo(ReqNum)
		eBody = sBody
		eBoolHtml=true
		call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

	rsPur.Close
	rsEmp.movenext
	Wend
	else
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in PurchaseTeam Panel.</b></font></center>"
		Response.end
	end if

	End Function

 %>
	<%
		Response.Redirect ("Purchase_QuotationList.asp")
	%>
<!--#include file="../includes/connection_close.asp"-->
