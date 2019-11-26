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
	Dim iloop
	iLoop = Request.Form("hdCount")


	For i = 1 to iLoop

	if request.form("action_" & i) <> "" then

		QtyAccepted = Request.Form("hdQtyAccepted_" & i)
		ItemDesc = Replace(Server.HTMLEncode(Request.Form("hdItemDesc_" & i)),"'","''")
		PurOrdNo = Request.Form("hdPurOrdNo_" & i)
		Remarks = Request.Form("Remarks_" & i)
		RequisitionNo = Request.Form("hdReqNo_" & i)
		GRNNo  = Request.form("hdGRNNo_" & i)

		sql = "update tbl_Psystem_GRN Set RemarksOnAccOrRej = '" & Replace(Server.HTMLEncode(Remarks),"'","''") & "', isAccepted = 2 where GRNNo = "& GRNNo &" and ItemDescription = '" & ItemDesc & "' and PurOrderNo = "& PurOrdNo &" and RequisitionId = "& RequisitionNo & " "
		'Response.write sql &  "<br>"
		Call DoSql(sql)

	End if

	Next

	Call fsSendMail_PurchaseTeam()

 %>
 <%
 	Private Function fsSendMail_PurchaseTeam()

	'-----Employer --------------
	EmployeeId=Session("Employee_Id")
	sql="sp_PSystem_GetLoggedEmployeeNameAndEmail '" & EmployeeId & "'"
	call RunSql(sql,rsEmp)
	eFromName=rsEmp("EmployeeName")
	eFromEmail=rsEmp("EmployeeEmail")
	rsEmp.close()
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
			" The synopses of the rejected GRN are as follows: " & _
			"</font><br><br>" & _
			"<table width='90%' border='0' cellspacing='2' cellpadding='2' align='center'>"
		  	sql = "Select distinct GRNNo,PurOrderNo,RequisitionId,PartyChallanNo,PartyChallanDate,SecurityEntryNo,DeliveryDate,LLRRNo,vehicleNo,SupplierName,Remarks from tbl_PSystem_GRN where GRNNo = "& GRNNo &" and isAccepted= 2"
			Call RunSql(sql,rsGRN)

			if rsGRN.EOF = false then
				ReqNo = rsGRN("RequisitionId")
				sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& ReqNo &" "
				call RunSql(sql,rsRec)
				if rsRec.Eof = false then
					ReqNum = rsRec("RequisitionNum")
				end if
				rsRec.Close
				
				PurOrdNo = rsGRN("PurOrderNo")
				sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrdNo &" "
				call RunSql(sql,rsPONum)
				if rsPONum.Eof = false then
					PurOrderNum = rsPONum("PurOrderNum")
				end if
				rsPONum.Close
				GRNNo = rsGRN("GRNNo")
				sql = " Select GRNNum  from tbl_Psystem_GRN where GRNNo = "& GRNNo &" "
				call RunSql(sql,rsGRNNum)
				if rsGRNNum.Eof = false then
					GRNNum = rsGRNNum("GRNNum")
				end if
				rsGRNNum.Close
				
				sSupName = rsGRN("SupplierName")
				sql= "Select SupplierAddress from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
				Call RunSql(sql,rsSup)
				if Not rsSup.Eof then
					sSupAddr = rsSup("SupplierAddress")
				end if
				rsSup.Close

       sBody = sBody &  "<tr bgcolor=#108ed6> <td>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>GRN No </b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase Order No</b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase Request No</b></font></div>" & _
               " </td> </tr> <tr bgcolor=#DFF2FC><td>" & _
               " <div align='center'> <font face='Trebuchet MS'> " & GetGRNNo(GRNNum) & "</font></div>" & _
               " </td> <td>" & _
               " <div align='center'> <font face='Trebuchet MS'>"& GetPurchaseOrderNo(PurOrderNum) &"</font></div>" & _
               " </td> <td>" & _
               " <div align='center'> <font face='Trebuchet MS'>"& GetPurchaseRequisitionNo(ReqNum) &"</font></div>" & _
               " </td> </tr> <tr bgcolor=#108ed6>  <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Supplier Info </b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Party Challan No</b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Party Challan Date</b></font></div>" & _
               " </td>  </tr>" & _
			   " <tr> <td rowspan='5' vAlign='top' bgcolor=#DFF2FC> " & _
               " <div align='center'></div>" & _
               " <div align='center'> <font face='Trebuchet MS'>" & sSupName & "<br>" & sSupAddr & "</font></div>" & _
               " </td><td bgcolor=#DFF2FC>" & _
               " <div align='center'> <font face='Trebuchet MS'>"& rsGRN("PartyChallanNo") &"</font></div>" & _
               " </td><td bgcolor=#DFF2FC>" & _
               " <div align='center'> <font face='Trebuchet MS'>"& SetDateFormat(rsGRN("PartyChallanDate")) &"</font></div>" & _
               " </td> </tr>" & _
               " <tr><td bgcolor=#108ed6>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Security Gate Entry No </b></font></div>" & _
               " </td><td bgcolor=#108ed6>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Delivery Date</b></font></div>" & _
               " </td></tr>" & _
               " <tr><td bgcolor=#DFF2FC>" & _
               " <div align='center'><font face='Trebuchet MS'> "& rsGRN("SecurityEntryNo") &"</font></div>" & _
               " </td> <td bgcolor=#DFF2FC>" & _
               " <div align='center'> <font face='Trebuchet MS'>"& SetDateFormat(rsGRN("DeliveryDate")) &"</font></div>" & _
               " </td> </tr>" & _
               " <tr>  <td bgcolor=#108ed6> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>LL/ RR No</b></font></div>" & _
               " </td> <td bgcolor=#108ed6>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Vehicle No</b></font></div>" & _
               " </td> </tr>" & _
               " <tr>  <td bgcolor=#DFF2FC> " & _
               " <div align='center'> <font face='Trebuchet MS'>"& rsGRN("LLRRNo") &"</font></div>" & _
               " </td> <td bgcolor=#DFF2FC>" & _
               "  <div align='center'> <font face='Trebuchet MS'>"& rsGRN("VehicleNo") &"</font></div>" & _
               " </td> </tr>"

			end if

    sBody = sBody &  " <tr><td colspan='3' vAlign='top'>" & _
               "  <table width='100%' border='0' cellpadding='2' cellspacing='2'>" & _
               "  <tr bgcolor=#108ed6>  <td> " & _
               "  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl.No</b></font></div>" & _
               "  </td><td> " & _
               "  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Item Description</b></font></div>" & _
               "  </td><td> " & _
               "  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Received</b></font></div>" & _
               "  </td><td> " & _
               "  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Accepted</b></font></div>" & _
               "  </td><td> " & _
               "  <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity Rejected</b></font></div>" & _
               "  </td> </tr>"

			sql = "Select ItemDescription,QtyReceived,QtyAccepted,QtyRejected from tbl_Psystem_GRN where GRNNo = "& GRNNo &" and isAccepted = 2"
			Call runSql(sql,rsInfo)
			i = 1
			While Not rsInfo.Eof

    sBody = sBody &  " <tr bgcolor=#DFF2FC> <td> " & _
                " <div align='center'> <font face='Trebuchet MS'>"& i &"</font></div> </td><td>" & _
                " <div align='center'> <font face='Trebuchet MS'>"& rsInfo("ItemDescription") &"</font></div></td> <td>" & _
                " <div align='center'> <font face='Trebuchet MS'>"& rsInfo("QtyReceived") &"</font></div></td><td>" & _
                " <div align='center'> <font face='Trebuchet MS'>"& rsInfo("QtyAccepted") &"</font></div></td><td>" & _
                " <div align='center'> <font face='Trebuchet MS'>"& rsInfo("QtyRejected") &"</font></div></td></tr>"

			i = i + 1
			rsInfo.movenext
			Wend
			rsInfo.Close

    sBody = sBody & " </table> </td></tr>" & _
             " <tr> " & _
             " <td colspan='3' bgcolor=#DFF2FC><font face='Trebuchet MS'><b>Remarks : </b> "& rsGRN("Remarks") &"</font> </td>" & _
		   	 "</tr><tr bgcolor=#108ed6>" &_
			 "<td colspan='3' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail.</font></td>" &_
			 "</tr>" &_
		     "</table>"

		eSubject = "Rejected GRN : " & GetGRNNo(GRNNum)
		eBody = sBody
		eBoolHtml=true
		Call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

	rsPur.Close
	rsEmp.movenext
	Wend
	else
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in PurchaseTeam Panel.</b></font></center>"
		Response.end
	end if


	End Function

 %>
 <script language="javascript">
 function redirect()
 {
 document.FinalForm.method="Post";
 document.FinalForm.action="GRNDetails_view.asp"
 document.FinalForm.submit();

 }
 </script>
<html>
<body onLoad="javascript:redirect()">
<form name="FinalForm">
<input type="hidden" name="hdGRNNo" value="<%=GRNNo%>">
</form>
</body></html>

<!--#include file="../includes/connection_close.asp"-->