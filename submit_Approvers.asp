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
	Dim iloop,i,Qty,isVelan,iStatus,QtyReq,QtyApp
	Dim AppQty,iCount,PrjId,sAction
	iLoop = Request.Form("hdCount")
	'Response.write ("count" & iLoop & "<br>")

	for i = 1 to iLoop

	'Response.write request.form("action_" & i)
	if request.form("action_" & i) <> "" then

		Qty = request.form("Quantity_" & i)
		if request.form("Ownby_" & i) = "isChecked" then
			isVelan = 1
		else
			isVelan = 0
		end if

		PrjID = request.form("hdPrjId_" & i)
		ReqID = request.form("hdReqID_" & i)
		IDesc = Replace(Server.HTMLEncode(request.form("hdIDesc_"& i)),"'","''")
		'Response.write "IDesc" & IDesc

		sql = sql_GetQtyRequired(ReqId,PrjId,IDesc)
		call RunSql(sql,rsQty)
		if rsQty.Eof = false then
			QtyReq = rsQty("QuantityRequested")
			QtyApp = rsQty("QuantityApproved")
		end if
		AppQty =  (cInt(QtyApp) + (Qty))
		if cInt(QtyReq) = AppQty then
			iStatus = 1
		else
			iStatus = 2
		end if

		sql = "update tbl_PSystem_PurchaseRequestTransaction set QuantityApproved = " & AppQty & ", IsVelankaniAsset = " & isVelan & ", Status = " & iStatus & " "&_
		   " where RequisitionId = " & ReqID & " and ItemDescription = '" & IDesc & "' and ProjectId = " & Prjid & " "

		Call DoSql(sql)

		sql = "sp_itbl_PSystem_TransactionDetails  " & ReqID & " ,'" & IDesc & "'," & Prjid & "," & Qty & "," & iStatus & " "
		'Response.write sql
		call Dosql(sql)

	end if
	next
	
	sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& ReqID &" "
	call RunSql(sql,rsRec)
	if rsRec.Eof = false then
		ReqNum = rsRec("RequisitionNum")
	end if
	rsRec.Close
	Call fsSendMail_Employee()
	Call fsSendMail_PurchaseTeam()

	%>

<%
	Private Function fsSendMail_Employee()
	'-------- Employeer ----------
	sql = "select EmployeeID from tbl_Psystem_PurchaseRequestMaster where RequisitionID = "& ReqID &" "
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
	efromEmail=rsApprover("EmployeeEmail")
	rsApprover.close()

	sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
			"</b><br>" & _
			" The synopses of the approved purchase request are as follows: " & _
			"</font><br><br>" & _
			"<table width='95%' border='0' align ='center' cellspacing='2' cellpadding='2'>" &_
			"<tr bgcolor=#108ed6>" &_
			"<td colspan='9' align='center'><font face='Trebuchet MS' color='#ffffff'><b>Approved Items</b></font></td>" &_
			"</tr><tr bgcolor=#108ed6>" &_
			"<td colspan='9' align='left'>&nbsp;<font face='Trebuchet MS' color='#ffffff'><b>Approvers Requisition No : </b> " & GetPurchaseRequisitionNo(ReqNum) & "</font></td>" &_
			"</tr>" &_
			"<tr bgcolor=#108ed6>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Item Description</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Project</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purpose</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Required</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Approved</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Required Date</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Possible Source</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Action</b></font></td>" &_
			"</tr>"

			sql= sql_GetRequiredItems(ReqId)
			call RunSql(sql,rsItems)
			iCount = 1
			while not rsItems.eof
				PrjId = rsItems("ProjectId")
				if rsItems("Status") = 1 then
					sAction = "Approved"
				elseif rsItems("Status") = 2 then
					sAction = "Partially Approved"
				else
					sAction = "Rejected"
				end if

				sql = sql_GetProjectName(PrjId)
				call RunSql(sql,rsPrj)
				if not rsPrj.EOF then
					sPrjName = rsPrj("ProjectName")
					rsPrj.Close
				end if

		sBody = sBody & "<tr bgcolor="& gsBGColorLight &">" &_
			"<td align='center'><font face='Trebuchet MS'> " & iCount & "</font></td>"  &_
			"<td><font face='Trebuchet MS'>" & rsItems("ItemDescription") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & sPrjName & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("Purpose") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("QuantityRequested") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("QuantityApproved") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & SetDateFormat(rsItems("RequiredDate")) & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("PossibleSource") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & sAction & " </font></td>" &_
			"</tr>"

			iCount=iCount+1
			rsItems.MoveNext
			Wend
			rsItems.Close
			rsReqester.Close
			rsEmp.Close

	sBody = sBody &	"</tr><tr bgcolor=#108ed6>" &_
			"<td colspan='9' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail.</font></td>" &_
			"</tr>" &_
		    "</table>"



	'Response.write sBody & "<br>"
	eSubject = "Approved Purchase Request Details : " & GetPurchaseRequisitionNo(ReqNum)
	eBody = sBody
	eBoolHTML = True
	call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

	End Function
 %>

 <%
	Private function fsSendMail_PurchaseTeam()

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
			" </b><br>" & _
			" The synopses of the approved purchase request are as follows: " & _
			"</font><br><br>" & _
			"<table width='90%' border='0' align ='center' cellspacing='2' cellpadding='2'>" &_
			"<tr bgcolor=#108ed6>" &_
			"<td colspan='7' align='center'><font face='Trebuchet MS' color='#ffffff'><b>Approved Items</b></font></td>" &_
			"</tr><tr>" &_
			"<td colspan='7' align='left'bgcolor=#108ed6>&nbsp;<font face='Trebuchet MS' color='#ffffff'><b>Approvers RequisitionNo : </b> " & GetPurchaseRequisitionNo(ReqNum) & "</font></td>" &_
			"</tr>" &_
			"<tr bgcolor=#108ed6>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Item Description</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Project</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purpose</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Approved</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Required Date</b></font></td>" &_
			"<td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Possible Source</b></font></td>" &_
			"</tr>"

			sql= "Select * from tbl_Psystem_PurchaseRequestTransaction where (Status = 1  or Status = 2 or Status =3) and RequisitionId = " & ReqId & " "
			call RunSql(sql,rsItems)
			iCount = 1
			while not rsItems.eof
				PrjId = rsItems("ProjectId")

				sql = sql_GetProjectName(PrjId)
				call RunSql(sql,rsPrj)
				if not rsPrj.EOF then
					sPrjName = rsPrj("ProjectName")
					rsPrj.Close
				end if

		sBody = sBody & "<tr bgcolor="& gsBGColorLight &">" &_
			"<td align='center'><font face='Trebuchet MS'>" & iCount & "</font></td>"  &_
			"<td><font face='Trebuchet MS'>" & rsItems("ItemDescription") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & sPrjName & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("Purpose") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("QuantityApproved") & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & SetDateFormat(rsItems("RequiredDate")) & "</font></td>" &_
			"<td><font face='Trebuchet MS'>" & rsItems("PossibleSource") & "</font></td>" &_
			"</tr>"

			iCount=iCount+1
			rsItems.MoveNext
			Wend
			rsItems.Close

		sBody = sBody &	"</tr><tr bgcolor=#108ed6>" &_
				"<td colspan='9' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail. </font></td>" &_
				"</tr>" &_
				"</table>"

		'Response.write sBody & "<br>"
		eSubject = "Approved Purchase Request Details : " & GetPurchaseRequisitionNo(ReqNum)
		eBody = sBody
		eBoolHTML = True
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
		Response.redirect ("PurchaseRequestList.asp")
	%>

<!--#include file="../includes/connection_close.asp"-->