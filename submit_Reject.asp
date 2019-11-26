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
<!--#include file="../includes/MailDesign.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
 <%
	Dim iloop,i,Qty,isVelan,iStatus
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
		iStatus = 3

		PrjID = request.form("hdPrjId_" & i)
		ReqID = request.form("hdReqID_" & i)
		IDesc = Replace(Server.HTMLEncode(request.form("hdIDesc_"& i)),"'","''")

		sql = sql_GetQtyRequired(ReqId,PrjId,IDesc)
		call RunSql(sql,rsQty)
		QtyReq = rsQty("QuantityRequested")
		QtyApp = rsQty("QuantityApproved")

		AppQty =  (cInt(QtyApp) + Qty)



		sql = "update tbl_PSystem_PurchaseRequestTransaction set  IsVelankaniAsset = " & isVelan & ", Status = " & iStatus & " "&_
			  " where RequisitionId = " & ReqID & " and ItemDescription = '" & IDesc & "' and ProjectId = " & Prjid & " "
		Call DoSql(sql)

		sql = "sp_itbl_PSystem_TransactionDetails  " & ReqID & " ,'" & IDesc & "'," & Prjid & "," & AppQty & "," & iStatus & " "
		'call Dosql(sql)

	end if
	next

	sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& ReqID &" "
	call RunSql(sql,rsRec)
	if rsRec.Eof = false then
		ReqNum = rsRec("RequisitionNum")
	end if
	rsRec.Close

	Call fsSendMail_Employee()

	Response.redirect ("PurchaseRequestList.asp")

 %>

 <%
	Private Function fsSendMail_Employee()

	Dim sBody

	'-----Employer --------------
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
	eFromEmail=rsApprover("EmployeeEmail")
	rsApprover.close()

 sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
		" </b><br>" & _
		" The synopses of the rejected purchase request are as follows: " & _
		"</font><br><br>" & _
		" <table width='100%' border='0' cellspacing='2' cellpadding='2'> " & _
		" <tr bgcolor=#108ed6>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>ItemDescription</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purpose</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Possible Source</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Special Instructions</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Required Date</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Project</b></font></div> </td>" & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Approx Cost</b></font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase / Service</b></font></div> </td>" & _
  		" </tr> "

  sql = "Select * from tbl_Psystem_PurchaseRequestTransaction  where status = 3 and RequisitionId = "& ReqID &" "
  Call RunSql(sql,rsRej)
  j = 1
  While Not rsRej.EOF

		Prj = rsRej("ProjectId")
		sql = sql_GetProjectName(Prj)
		call RunSql(sql,rsPrj)
		if not rsPrj.EOF then
			sPrjName = rsPrj("ProjectName")
			rsPrj.Close
		end if

		if rsRej("RupeeOrDollar") = 0 then
			sCurr = "Rupee(s)"
		else
			sCurr = "Dollar(s)"
		end if

		if rsRej("PurchaseOrService") = 0 then
			sPurSer = "Purchase"
		else
			sPurSer = "Service"
		end if

	sBody = sBody & "<tr bgcolor=#DFF2FC> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & j & "</font></div> </td> " & _
    	" <td> <div align='left'><font face='Trebuchet MS'>" & rsRej("ItemDescription") & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>"  & rsRej("Purpose") & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & rsRej("QuantityRequested") & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & rsRej("PossibleSource") & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & rsRej("SpecialInstruction") & "</font></div> </td>" & _
   		" <td> <div align='center'><font face='Trebuchet MS'>" & SetDateFormat(rsRej("RequiredDate")) & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & sPrjName & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & rsRej("ApproxUnitCost") & " " & sCurr & "</font></div> </td> " & _
    	" <td> <div align='center'><font face='Trebuchet MS'>" & sPurSer & "</font> </div> </td> " & _
  		" </tr> "
	  j = j + 1
	  rsRej.Movenext
	  Wend
	  rsRej.Close
 	  rsReqester.Close
	  rsEmp.Close

	sBody = sBody &	"</tr><tr bgcolor=#108ed6>" &_
			"<td colspan='10' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail. </font></td>" &_
			"</tr>" &_
		    "</table>"

	'Response.write sBody
	eSubject = "Rejected Purchase Request : " & GetPurchaseRequisitionNo(ReqNum)
	eBody = sBody
	eBoolHTML = True
	Call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

	End Function

 %>
<!--#include file="../includes/connection_close.asp"-->