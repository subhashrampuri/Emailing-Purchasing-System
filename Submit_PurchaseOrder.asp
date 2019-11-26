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
	ReqId = request.form("hdReqId")
	RequiredDate =  request.form("RequiredDate")
	Tax = request.form("Tax")
	Warranty = request.form("Warranty")
	PTerms = request.form("PaymentTerms")
	Others=  request.form("Others")
	GTotal = Request.Form("hdGTotal")
	Supplier = Request.Form("hdSupplier")
	SupplierAddr = Request.Form("hdSupAddr")
	PurOrderNo = flGeneratePONo()

	sql = " select Counter+1 as Counter from tbl_PSystem_Control where DateDiff(DD,getDate(),EndDate)>=0 and ForType = 'PO' "
	call RunSql(sql,rsPONum)
	PurOrderNum = rsPONum(0)

	sql = "sp_itbl_PSystem_PurchaseOrder "& PurOrderNo &"," & ReqId & ", '" & RequiredDate & "','" & Replace(Server.HTMLEncode(PTerms),"'","''") & "', '" & Replace(Server.HTMLEncode(Others),"'","''") & "', " & GTotal & ",'" & Date() & "', "& PurOrderNum &" "
	'Response.write sql
	call Dosql(sql)

	sql = "Update tbl_Psystem_Control set Counter = "& PurOrderNum &" where DateDiff(DD,getDate(),EndDate)>=0 and ForType = 'PO' "
	Call Dosql(sql)

	sql = "Update tbl_Psystem_Quotations Set isApproved = 4, PurOrderNo = "& PurOrderNo &" , PurOrderDate = '" & Date() & "' where RequisitionId = " & ReqId & " and isApproved = 1 and SupplierName = '" & Supplier & "' "
	'Response.write sql
	call Dosql(sql)


	Call fsSendMail_Employee()
	Call fsSendMail_FinanceTeam()
 %>
 	<%
		Private Function flGeneratePONo()
		Dim lPrimaryKey

		sql = "SELECT max(PurOrderNo) FROM tbl_Psystem_PurchaseOrder"
		call RunSql(sql,rsPO)

		if isNull(rsPO(0)) then
			lPrimaryKey=0
		else
			lPrimaryKey=rsPO(0)
		end if

		Set rsPO = Nothing
		flGeneratePONo=CLNG(lPrimaryKey)+1
	End Function

	%>

	<%
	Private Function fsSendMail_Employee()

	Dim sBody
	'-----Employer --------------
	sql = "select EmployeeID from tbl_Psystem_PurchaseRequestMaster where RequisitionID = "& ReqId &" "
	Call RunSql(sql,rsEmp)

	if rsEmp.Eof = false then
		EmpID = rsEmp("EmployeeID")
	end if

	str = "select dbo.fn_TSystem_EmployeeName('" & EmpID & "') as Name, dbo.fn_TSystem_EmployeeEmail('" & EmpID & "') as Email"
	Call RunSql(str,rsReqester)

	if rsReqester.EOF = false then
		eToName = rsReqester("Name")
		eToEmail = rsReqester("Email")
	end if
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
	'-------------------------------------

	sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
		"</b><br>" & _
		" The synopses of the purchase order is as follows: " & _
		"</font><br><br>" & _
		" <table width='100%' border='0' cellspacing='2' cellpadding='2'>" & _
		" <tr bgcolor=#108ed6> " &_
		" <td colspan='6'> " &_
		" <div align='center' ><font face='Trebuchet MS' color='#ffffff'><b>Purchase Order Released</b></font></div>" & _
		" </td>  </tr>" & _
		" <tr bgcolor=#108ed6> " & _
		" <td colspan='2'> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Supplier Info</b></font></div>" & _
		" </td> <td colspan='2'> " & _
		" <div align='center' nowrap><font face='Trebuchet MS' color='#ffffff'><b>Purchase Order No </b></font></div>" & _
		" </td><td colspan='2' align='center'> <font face='Trebuchet MS' color='#ffffff' ><b>Date</b></font></td> </tr>" & _
		" <tr> <td colspan='2' bgcolor=#DFF2FC> " & _
		" <div align='center'> <font face='Trebuchet MS' >" & Supplier & "<br>" & SupplierAddr & " </font></div>" & _
		" </td>	<td colspan='2'bgcolor=#DFF2FC> " & _
		" <div align='center'><font face='Trebuchet MS' >" & GetPurchaseOrderNo(PurOrderNum) & "</font></div>" & _
		" </td><td colspan='2' bgcolor=#DFF2FC align='center'> <font face='Trebuchet MS' >" & SetDateFormat(Date()) & " </font></td>" & _
		" </tr> <tr bgcolor=#108ed6> <td> " & _
		"  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>ItemDescription</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff' ><b>Tax Percent</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Price</b></font></div>" & _
		" </td>	<td> " & _
	  	" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Amount</b></font></div>" & _
		" </td>  </tr>"

		 sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and SupplierName = '" & Supplier & "' and RequisitionId = "& ReqId &" "
 		 call RunSql(sql,rsInfo)
		 i = 1
		 while not rsInfo.EOF
		 if rsInfo("Currency") = -1 then
		 	Curr = "Rs."
		 else
		 	Curr = "$"
		 end if
		 Qty = rsInfo("Quantity")
		 Price = rsInfo("UnitPrice")
		 TaxPercent = rsInfo("TaxPercent")
		 Tax = (cDbl(TaxPercent) / 100)
		 Amount = (cInt(Qty) * cDbl(Price))
		 if rsInfo("isTaxIncludedOrExcluded") = -1 then
		 	Total = ((Amount) + ((Amount)* cDbl(Tax)))
		 else
		 	Total = Amount
		 end if


		sBody = sBody & " <tr bgcolor=#DFF2FC>" & _
			" <td><div align='center'><font face='Trebuchet MS'>" &  i & "</font></div></td>" & _
			" <td><div align='center'><font face='Trebuchet MS'>" & rsInfo("ItemDescription") & "</font></div> </td>" & _
			" <td><div align='center'><font face='Trebuchet MS'>" & rsInfo("TaxPercent") & " " & "%"  & "</font></div> </td>" & _
			" <td><div align='center'><font face='Trebuchet MS'>" & rsInfo("Quantity") & " </font></div></td>" & _
			" <td><div align='center'><font face='Trebuchet MS'>" & Curr & " " & rsInfo("UnitPrice") & "</font></div></td>" & _
			" <td><div align='center'><font face='Trebuchet MS'>" & FormatNumber(Total,2) & "</font></div></td>" & _
		  	"</tr>"

			  i = i + 1
			  ETotal = cDbl(ETotal) + cDbl(FormatNumber(Total,2))
			  rsInfo.movenext
			  Wend
			  rsInfo.Close
			rsReqester.Close
			rsEmp.Close

		  sBody = sBody & "<tr> <td colspan='4'>&nbsp;</td>" & _
			" <td bgcolor=#DFF2FC>" & _
			"  <div align='right'><font face='Trebuchet MS' ><b>Grand Total: </b></font></div></td>" & _
			" <td bgcolor=#DFF2FC>" & _
			" <div align='center'><font face='Trebuchet MS'>" & Curr & " " & FormatNumber(ETotal,2) & "</font></div></td>" & _
		  	" </tr>" & _
			" <tr bgcolor=#108ed6><td colspan='6' align='left'> " &_
			"  <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail.</font>" & _
			" </td></tr></table>"

		eSubject = "Released Purchase Order: " & GetPurchaseOrderNo(PurOrderNum)
		eBody = sBody
		eBoolHtml=true
		call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)
	End Function
	%>

<%
	Private Function fsSendMail_FinanceTeam()

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
	'-------Finance Manager-------------
	sql = "Select FinanceManager from tbl_Psystem_FinanceManager"
	call RunSql(sql,rsFM)

	if rsFM.EOF = false then
		While NOT rsFM.EOF
			EmpID = rsFM("FinanceManager")

			sql = "select dbo.fn_TSystem_EmployeeName('" & EmpID & "') As EmpName, dbo.fn_TSystem_EmployeeEmail('" & EmpID & "') as EmpEmail"
			call RunSql(sql,rsFMDetails)
			if rsFMDetails.EOf = false then
				eToName = rsFMDetails("EmpName")
				eToEMail = rsFMDetails("EmpEmail")
			end if


	sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
		"</b><br>" & _
		" The synopses of the purchase order is as follows: " & _
		"</font><br><br>" & _
		" <table width='100%' border='0' cellspacing='2' cellpadding='2'>" & _
		" <tr bgcolor=#108ed6> " &_
		" <td colspan='6'> " &_
		" <div align='center' ><font face='Trebuchet MS' color='#ffffff'><b>Purchase Order Released</b></font></div>" & _
		" </td>  </tr>" & _
		" <tr bgcolor=#108ed6> " & _
		" <td colspan='2'> " & _
		" <div align='center' ><font face='Trebuchet MS' color='#ffffff'><b>Supplier Info</b></font></div>" & _
		" </td> <td colspan='2'> " & _
		" <div align='center' nowrap><font face='Trebuchet MS' color='#ffffff'><b>Purchase Order No </b></font></div>" & _
		" </td><td colspan='2' align='center'> <font face='Trebuchet MS' color='#ffffff'><b>Date</b></font></td> </tr>" & _
		" <tr> <td colspan='2' bgcolor=#DFF2FC> " & _
		" <div align='center'> <font face='Trebuchet MS'>" & Supplier & "<br>" & SupplierAddr & "</font></div>" & _
		" </td>	<td colspan='2'bgcolor=#DFF2FC> " & _
		" <div align='center'><font face='Trebuchet MS' >" & GetPurchaseOrderNo(PurOrderNum) & "</font></div>" & _
		" </td><td colspan='2' bgcolor=#DFF2FC align='center'> <font face='Trebuchet MS'>" & SetDateFormat(Date()) & " </font></td>" & _
		" </tr> <tr bgcolor=#108ed6> <td> " & _
		"  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>ItemDescription</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff' ><b>Tax Percent</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity</b></font></div>" & _
		" </td>	<td> " & _
		" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Price</b></font></div>" & _
		" </td>	<td> " & _
	  	" <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Amount</b></font></div>" & _
		" </td>  </tr>"

		 sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and SupplierName = '" & Supplier & "' and RequisitionId = "& ReqId &" "
 		 call RunSql(sql,rsInfo)
		 i = 1
		 while not rsInfo.EOF
		 if rsInfo("Currency") = -1 then
		 	Curr = "Rs."
		 else
		 	Curr = "$"
		 end if
		 Qty = rsInfo("Quantity")
		 Price = rsInfo("UnitPrice")
		 TaxPercent = rsInfo("TaxPercent")
		 Tax = (cDbl(TaxPercent) / 100)
		 Amount = (cInt(Qty) * cDbl(Price))
		 if rsInfo("isTaxIncludedOrExcluded") = -1 then
		 	Total = ((Amount) + ((Amount)* cDbl(Tax)))
		 else
		 	Total = Amount
		 end if


		sBody = sBody & " <tr bgcolor=#DFF2FC>" & _
			" <td><div align='center'><font face='Trebuchet MS' >" &  i & "</font></div></td>" & _
			" <td><div align='center'><font face='Trebuchet MS' >" & rsInfo("ItemDescription") & "</font></div> </td>" & _
			" <td><div align='center'><font face='Trebuchet MS' >" & rsInfo("TaxPercent") & " " & "%"  & "</font></div> </td>" & _
			" <td><div align='center'><font face='Trebuchet MS' >" & rsInfo("Quantity") & " </font></div></td>" & _
			" <td><div align='center'><font face='Trebuchet MS' >" & Curr & " " & rsInfo("UnitPrice") & "</font></div></td>" & _
			" <td><div align='center'><font face='Trebuchet MS' >" & FormatNumber(Total,2) & "</font></div></td>" & _
		  	"</tr>"

			  i = i + 1
			  FTotal = cDbl(FTotal) + cDbl(FormatNumber(Total,2))
			  rsInfo.movenext
			  Wend
			  rsInfo.Close

		  sBody = sBody & "<tr> <td colspan='4'>&nbsp;</td>" & _
			" <td bgcolor=#DFF2FC>" & _
			"  <div align='right'><font face='Trebuchet MS' ><b>Grand Total: </b></font></div></td>" & _
			" <td bgcolor=#DFF2FC>" & _
			" <div align='center'><font face='Trebuchet MS' >" & Curr & " " & FormatNumber(FTotal,2) & "</font></div></td>" & _
		  	" </tr>" & _
			" <tr bgcolor=#108ed6><td colspan='6' align='left'> " &_
			"  <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail.</font>" & _
			" </td></tr></table>"

			'Response.write sBody

		eSubject = "Released Purchase Order: " & GetPurchaseOrderNo(PurOrderNum)
		eBody = sBody
		eBoolHtml=true

		call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

		rsFMDetails.close
		rsFM.movenext
		Wend
		else
			Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in Finance Manager Panel.</b></font></center>"
			Response.end
		end if

	End Function
	%>

	<script language="javascript">
	function redirect()
		{
			document.FinalForm.method="Post";
			document.FinalForm.action="AckPurOrderReleased.asp"
			document.FinalForm.submit();
		}
	</script>
	<html>
	<body onLoad="javascript:redirect();">
	<form name="FinalForm">
		<input type="hidden" name="hdReqId" value="<%=ReqId%>">
		<input type="hidden" name="hdPurOrdNo" value="<%=PurOrderNo%>">
 	    <input type="hidden" name="hdSupplier" value="<%=Supplier%>">
		<input type="hidden" name="hdSupAddr" value = "<%=SupplierAddr%>">
	</form>
	</body>
	</html>


<!--#include file="../includes/connection_close.asp"-->
