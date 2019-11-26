<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Subhash Rampuri
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<!--#include file="../includes/mail.asp"-->
<%
	Sub SendRequestToApprover(EmployeeEmail,ApproverEmail,MailSubject,MailBody)
		eFromEmail=EmployeeEmail
		eToName = EmployeeName
		eToEmail=ApproverEmail
		eSubject=MailSubject
		eBody=MailBody
	'	Response.write eBody
		eBoolHtml=true
		call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)
	End Sub
%>

<%
	EmployeeId=Session("Employee_Id")
	sql="select dbo.fn_PSystem_EmployeeDepartmentName('" & EmployeeId & "')"
	call RunSql(sql,rsEmployeeDepartmentName)
	EmployeeDepartmentName=rsEmployeeDepartmentName(0)
	rsEmployeeDepartmentName.close()
	sql="sp_PSystem_GetActiveApprover"
	call RunSql(sql,rsApprover)
	if rsApprover.eof then
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in Approvers Panel.</b></font></center>"
		Response.end
	end if
	ApproverId=rsApprover("EmployeeId")
	ApproverEmail=rsApprover("EmployeeEmail")
	rsApprover.close()
	sql="sp_PSystem_GetLoggedEmployeeNameAndEmail '" & EmployeeId & "'"
	call RunSql(sql,rsEmp)
	EmployeeName=rsEmp("EmployeeName")
	EmployeeEmail=rsEmp("EmployeeEmail")
	rsEmp.close()
	sql="sp_PSystem_AddPurchaseRequest '" & EmployeeId & "','" & ApproverId & "'"
	call DoSql(sql)
	sql="select @@Identity"
	'sql = "select Counter from tbl_PSystem_Control where DateDiff(DD,getDate(),EndDate)>=0 and ForType = 'PR' "
	call RunSql(sql,rsIdentity)
	RequisitionId=rsIdentity(0)
	rsIdentity.close()
	arrItems=split(Request.Form("ItemList"),",")
	for i=0 to ubound(arrItems)
		ProjectId=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		ItemDescription=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		ProjectName=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		Purpose=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		QuantityRequested=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		PurchaseOrService=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		RequiredDate=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		if arrItems(i)="-" then
			ApproxUnitCost=0
		else
			arrTemp=split(arrItems(i)," ")
			ApproxUnitCost=arrTemp(0)
		end if
		i=i+1
		PossibleSource=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		SpecialInstruction=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		ServiceType=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))
		i=i+1
		CCurrency=Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))

		'sql="sp_PSystem_AddPurchaseRequestItem " & RequisitionId & ",'" & Replace(Server.HTMLEncode(ItemDescription),"'","''") & "'," & ProjectId & ",'" & Replace(Server.HTMLEncode(Purpose),"'","''") & "'," & QuantityRequested & ",'" & RequiredDate & "'," & ApproxUnitCost & "," & CCurrency & "," & ServiceType & ",'" & Replace(Server.HTMLEncode(PossibleSource),"'","''") & "','" & Replace(Server.HTMLEncode(SpecialInstruction),"'","''") & "'"
		'Response.write sql

		sql="sp_PSystem_AddPurchaseRequestItem " & RequisitionId & ",'" & ItemDescription & "'," & ProjectId & ",'" & Purpose & "'," & QuantityRequested & ",'" & RequiredDate & "'," & ApproxUnitCost & "," & CCurrency & "," & ServiceType & ",'" & PossibleSource & "','" & SpecialInstruction & "'"
		call DoSql(sql)

	next

	Dim sMailBody
	sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& RequisitionId &" "
	call RunSql(sql,rsRec)
	if rsRec.Eof = false then
		ReqNum = rsRec("RequisitionNum")
	end if
	rsRec.close
	sMailBody = fsMail()
	call SendRequestToApprover(EmployeeEmail,ApproverEmail,"Purchase Request for approval: " & GetPurchaseRequisitionNo(ReqNum),sMailBody)

	Call fsSendMail_ApproverTeam()
%>

<%

	Private function fsMail()
		sBody  = " <font face='Trebuchet MS'>Dear <b>" &  EmployeeName &", " &_
				"</b><br>" & _
				" The synopses of the purchase request is as follows: " & _
				"</font><br><br>" & _
				"<table width='98%' align='center' valign='top' cellspacing='2' cellpadding='2' border='0'>" &_
				"<TR height='25' bgcolor=#108ed6>" &_
				"<td colspan='10' align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase Request</b></font></td>" &_
				"</td>" &_
				"<tr height='25' bgcolor=#108ed6>" &_
				"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Employee : " & EmployeeName & " ( " & EmployeeId & " ) </font></td>" &_
				"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Department : " & EmployeeDepartmentName & " </font></td>" &_
				"</tr>" &_
				"<tr height='25' bgcolor=#108ed6>" &_
				"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Purchase Requisition No : " & GetPurchaseRequisitionNo(ReqNum) & " </font></td>" &_
				"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Requisition Date: " & SetDateFormat(Date()) & " </font></td>" &_
				"</tr>" &_
				"<tr height='25' bgcolor=#108ed6>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl. No.</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Item Description</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Project</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purpose</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Required</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Request Type</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Required Date</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Approx Unit Cost</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Possible Source</b></font></td>" &_
				"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Special Instructions</b></font></td>"

					lclstr_bgColor = gsBGColorLight
					sql="sp_PSystem_GetItemsByPurchaseRequisitionId '" & RequisitionId & "'"
					call RunSql(sql,rsItems)
					counter=1
					while not rsItems.eof
						if lclstr_bgColor = gsBGColorLight then
							lclstr_bgColor = gsBGColorDark
						else
							lclstr_bgColor = gsBGColorLight
						end if
			sBody = sBody & "<tr height='25' bgcolor=#DFF2FC>" &_
				"<td align='center'><font face='Trebuchet MS'> " & counter & "</font></td>"
					counter=counter+1
			sBody = sBody & "</td>"	 &_
				"<td> <font face='Trebuchet MS'>" & rsItems("ItemDescription") & "</font></td>" &_
				"<td> <font face='Trebuchet MS'>" & rsItems("Project") & "</font></td>" &_
				"<td> <font face='Trebuchet MS'>" & rsItems("Purpose") & "</font></td>" &_
				"<td> <font face='Trebuchet MS'>" & rsItems("QuantityRequested") & "</font></td>" &_
				"<td> <font face='Trebuchet MS'>" & rsItems("ServiceType") & "</font></td>" &_
				"<td> <font face='Trebuchet MS'>" & rsItems("RequiredDate") & "</font></td>"

				if rsItems("ApproxUnitCost")= "" then
					sBody = sBody	 & "<td> "-" </td>"
				else
					sBody = sBody & "<td> <font face='Trebuchet MS'>" &  rsItems("ApproxUnitCost") & " " & rsItems("Currency") & " </font></td>"
				end if

			sBody = sBody &	"<td> <font face='Trebuchet MS'>" & rsItems("PossibleSource") & "</font></td>"  &_
				"<td> <font face='Trebuchet MS'>" & rsItems("SpecialInstruction") & "</font></td>" &_
				"</tr>"
					rsItems.movenext
					wend
					rsItems.close()
		sBody = sBody &	"</tr><tr bgcolor=#108ed6 height='25'>" &_
				"<td colspan='10' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail. </font></td>" &_
				"</tr>" &_
		    	"</table>"

		fsMail = sBody

	End Function
%>
<%
	Private Function fsSendMail_ApproverTeam()
	'********** Employer Email**********
	sql="sp_PSystem_GetLoggedEmployeeNameAndEmail '" & EmployeeId & "'"
	call RunSql(sql,rsEmp)
	eFromName=rsEmp("EmployeeName")
	eFromEmail=rsEmp("EmployeeEmail")
	rsEmp.close()

	'***********Active Approver **********
	sql="sp_PSystem_GetActiveApprover"
	call RunSql(sql,rsApprover)
	if rsApprover.eof then
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in Approvers Panel.</b></font></center>"
		Response.end
	end if
	eToName=rsApprover("EmployeeName")
	eToEmail=rsApprover("EmployeeEmail")
	rsApprover.close()

		sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
			"</b><br>" & _
			" The synopses of the purchase request is as follows: " & _
			"</font><br><br>" & _
			"<table width='98%' align='center' valign='top' cellspacing='2' cellpadding='2' border='0'>" &_
			"<TR height='25' bgcolor=#108ed6>" &_
			"<td colspan='10' align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase Request: Acknowledgement</b></font></td>" &_
			"</td>" &_
			"<tr height='25' bgcolor=#108ed6>" &_
			"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Employee : " & EmployeeName & " ( " & EmployeeId & " ) </font></td>" &_
			"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Department : " & EmployeeDepartmentName & " </font></td>" &_
			"</tr>" &_
			"<tr height='25' bgcolor=#108ed6>" &_
			"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Purchase Requisition No : " & GetPurchaseRequisitionNo(ReqNum) & " </font></td>" &_
			"<td colspan='5'>&nbsp;<font face='Trebuchet MS' color='#ffffff'>Requisition Date: " & SetDateFormat(Date()) & " </font></td>" &_
			"</tr>" &_
			"<tr height='25' bgcolor=#108ed6>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl. No.</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Item Description</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Project</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purpose</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Required</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Request Type</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Required Date</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Approx Unit Cost</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Possible Source</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Special Instructions</b></font></td>"

				lclstr_bgColor = gsBGColorLight
				sql="sp_PSystem_GetItemsByPurchaseRequisitionId '" & RequisitionId & "'"
				call RunSql(sql,rsItems)
				counter=1
				while not rsItems.eof
					if lclstr_bgColor = gsBGColorLight then
						lclstr_bgColor = gsBGColorDark
					else
						lclstr_bgColor = gsBGColorLight
					end if
		sBody = sBody & "<tr height='25' bgcolor=#DFF2FC>" &_
			"<td align='center'><font face='Trebuchet MS'> " & counter & "</font></td>"
				counter=counter+1
		sBody = sBody & "</td>"	 &_
			"<td> <font face='Trebuchet MS'>" & rsItems("ItemDescription") & "</font></td>" &_
			"<td> <font face='Trebuchet MS'>" & rsItems("Project") & "</font></td>" &_
			"<td> <font face='Trebuchet MS'>" & rsItems("Purpose") & "</font></td>" &_
			"<td> <font face='Trebuchet MS'>" & rsItems("QuantityRequested") & "</font></td>" &_
			"<td> <font face='Trebuchet MS'>" & rsItems("ServiceType") & "</font></td>" &_
			"<td> <font face='Trebuchet MS'>" & rsItems("RequiredDate") & "</font></td>"

			if rsItems("ApproxUnitCost")= "" then
				sBody = sBody	 & "<td> "-" </td>"
			else
				sBody = sBody & "<td> <font face='Trebuchet MS'>" &  rsItems("ApproxUnitCost") & " " & rsItems("Currency") & " </font></td>"
			end if

		sBody = sBody &	"<td> <font face='Trebuchet MS'>" & rsItems("PossibleSource") & "</font></td>"  &_
			"<td> <font face='Trebuchet MS'>" & rsItems("SpecialInstruction") & "</font></td>" &_
			"</tr>"
				rsItems.movenext
				wend
				rsItems.close()
	sBody = sBody &	"</tr><tr bgcolor=#108ed6 height='25'>" &_
			"<td colspan='10' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail. </font></td>" &_
			"</tr>" &_
			"</table>"

		'Response.write sBody
		eSubject = "Purchase Request Acknowledgement: " & GetPurchaseRequisitionNo(ReqNum)
		eBody = sBody
		eBoolHtml=true
		call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

	End Function

%>
<form name="PostForm">
<input type="hidden" name="hdReqId" value="<%=RequisitionId%>">
</form>
<script language="javascript">
	var Req = document.PostForm.hdReqId.value;
	document.PostForm.method="Post"
	document.PostForm.action = "AckPurchaseRequest.asp"
	document.PostForm.submit();

</script>


<!--#include file="../includes/main_page_close.asp"-->