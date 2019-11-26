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

	Dim RequisitionNo,iLoop,i
	RequisitionNo = Request.Form("hdReqId")
	PurOrdNo = Request.Form("hdPurOrdNo")
	iLoop = Request.Form("hdCount")
	PartyChallanNo = Request.form("PartyChallanNo")
	PartyChallanDate = Request.Form("PartyChallanDate")
	SecutiryEntry = Request.Form("SecutiryEntry")
	DeliveryDate = Request.Form("DeliveryDate")
	LLRRNo = Request.Form("LLRRNo")
	VehicleNo = Request.Form("VehicleNo")
	Remarks = Request.Form("Remarks")
	GRNNo = flGenerateGRNNo()

	for i=1 to iLoop
	if request.form("action_" & i) <> "" then

		QtyReceived = request.form("QtyReceived_" & i)
		QtyAccepted = request.form("QtyAccepted_" & i)
		QtyRejected = request.form("QtyRejected_" & i)

		ItemDesc = request.form("hdItemDesc_" & i)
		ProjectID = request.form("hdPrjID_" & i)
		SupplierName = request.form("hdSupplier_" & i)

		sql = " select Counter+1 as Counter from tbl_PSystem_Control where DateDiff(DD,getDate(),EndDate)>=0 and ForType = 'GRN' "
		call RunSql(sql,rsGRNNum)
		GRNNum = rsGRNNum(0)

		sql = "sp_itbl_PSystem_GRN  "& GRNNo &","& PurOrdNo &","& RequisitionNo &", '" & Replace(Server.HTMLEncode(PartyChallanNo),"'","''") & "', '" &  PartyChallanDate & "', '" & Replace(Server.HTMLEncode(SecutiryEntry),"'","''") & "', '" & DeliveryDate & "', '" & Replace(Server.HTMLEncode(LLRRNo),"'","''") & "', '" & Replace(Server.HTMLEncode(VehicleNo),"'","''") & "', '" & Replace(Server.HTMLEncode(SupplierName),"'","''") & "', '" & Replace(Server.HTMLEncode(ItemDesc),"'","''") & "', "& QtyReceived &", "& QtyAccepted &", "& QtyRejected &", '" & Replace(Server.HTMLEncode(Remarks),"'","''") & "', "& GRNNum &" "
		'response.write sql
		call DoSql(sql)

		sql = "Update tbl_Psystem_Control set Counter = "& GRNNum &" where DateDiff(DD,getDate(),EndDate)>=0 and ForType = 'GRN' "
		Call Dosql(sql)

	end if
	Next
	Call fsSendMail_Employee()
%>
<%
	Private Function flGenerateGRNNo()
	Dim lPrimaryKey

	sql = "SELECT max(GRNNo) FROM tbl_Psystem_GRN"
	call RunSql(sql,rsGRN)

	if isNull(rsGRN(0)) then
		lPrimaryKey=0
	else
		lPrimaryKey=rsGRN(0)
	end if

		Set rsGRN = Nothing
		flGenerateGRNNo=CLNG(lPrimaryKey)+1
	End Function

%>
<%
	Private Function fsSendMail_Employee()
	'-----Logged Employee------------
	sql = "select EmployeeID from tbl_Psystem_PurchaseRequestMaster where RequisitionID = "& RequisitionNo &" "
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
	'------Active Purchaser Team------
	sql= "sp_PSystem_GetActivePurchaseTeam"
	call RunSql(sql,rsPur)
	if rsPur.eof then
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in PurchaseTeam Panel.</b></font></center>"
		Response.end
	end if
	eFromName = rsPur("EmployeeName")
	eFromEmail = rsPur("EmployeeEmail")
	rsPur.Close

	sBody  = " <font face='Trebuchet MS'>Dear <b>" &  eToName &", " &_
			"</b><br>" & _
			" The synopses of the GRN entry are as follows: " & _
			"</font><br><br>" & _
			"<table width='90%' border='0' cellspacing='2' cellpadding='2' align='center'>"

			sql = "Select distinct GRNNo,PurOrderNo,RequisitionId,PartyChallanNo,PartyChallanDate,SecurityEntryNo,DeliveryDate,LLRRNo,vehicleNo,SupplierName,Remarks from tbl_PSystem_GRN where GRNNo = "& GRNNo &" "
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
				sSupName = rsGRN("SupplierName")
				sql= "Select SupplierAddress from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
				Call RunSql(sql,rsSup)
				if Not rsSup.Eof then
					sSupAddr = rsSup("SupplierAddress")
				else
					sSupAddr = " "
				end if
				rsSup.Close

       sBody = sBody &  "<tr bgcolor=#108ed6> <td>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>GRN No </b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase Order No</b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Purchase Request No</b></font></div>" & _
               " </td> </tr> <tr bgcolor=#DFF2FC><td>" & _
               " <div align='center'><font face='Trebuchet MS' >" & GetGRNNo(GRNNum) & "</font></div>" & _
               " </td> <td>" & _
               " <div align='center'><font face='Trebuchet MS' >"& GetPurchaseOrderNo(PurOrderNum) &"</font></div>" & _
               " </td> <td>" & _
               " <div align='center'><font face='Trebuchet MS' >"& GetPurchaseRequisitionNo(ReqNum) &"</font></div>" & _
               " </td> </tr> <tr bgcolor=#108ed6>  <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Supplier Info </b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Party Challan No</b></font></div>" & _
               " </td> <td> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Party Challan Date</b></font></div>" & _
               " </td>  </tr>" & _
			   " <tr> <td rowspan='5' vAlign='top' bgcolor=#DFF2FC> " & _
               " <div align='center'></div>" & _
               " <div align='center'><font face='Trebuchet MS' >" & sSupName & "<br>" & sSupAddr & "</font></div>" & _
               " </td><td bgcolor=#DFF2FC>" & _
               " <div align='center'><font face='Trebuchet MS' >"& rsGRN("PartyChallanNo") &"</font></div>" & _
               " </td><td bgcolor=#DFF2FC>" & _
               " <div align='center'><font face='Trebuchet MS' >"& SetDateFormat(rsGRN("PartyChallanDate")) &"</font></div>" & _
               " </td> </tr>" & _
               " <tr><td bgcolor=#108ed6>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Security Gate Entry No </b></font></div>" & _
               " </td><td bgcolor=#108ed6>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Delivery Date</b></font></div>" & _
               " </td></tr>" & _
               " <tr><td bgcolor=#DFF2FC>" & _
               " <div align='center'><font face='Trebuchet MS' >"& rsGRN("SecurityEntryNo") &"</font></div>" & _
               " </td> <td bgcolor=#DFF2FC>" & _
               " <div align='center'><font face='Trebuchet MS' >"& SetDateFormat(rsGRN("DeliveryDate")) &"</font></div>" & _
               " </td> </tr>" & _
               " <tr>  <td bgcolor=#108ed6> " & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>LL/ RR No</b></font></div>" & _
               " </td> <td bgcolor=#108ed6>" & _
               " <div align='center'><font face='Trebuchet MS' color='#ffffff'><b>Vehicle No</b></font></div>" & _
               " </td> </tr>" & _
               " <tr>  <td bgcolor=#DFF2FC> " & _
               " <div align='center'><font face='Trebuchet MS' >"& rsGRN("LLRRNo") &"</font></div>" & _
               " </td> <td bgcolor=#DFF2FC>" & _
               "  <div align='center'><font face='Trebuchet MS' >"& rsGRN("VehicleNo") &"</font></div>" & _
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

			sql = "Select ItemDescription,QtyReceived,QtyAccepted,QtyRejected from tbl_Psystem_GRN where GRNNo = "& GRNNo &" "
			Call runSql(sql,rsInfo)
			i = 1
			While Not rsInfo.Eof

    sBody = sBody &  " <tr bgcolor=#DFF2FC> <td> " & _
                " <div align='center'><font face='Trebuchet MS' >"& i &"</font></div> </td><td>" & _
                " <div align='center'><font face='Trebuchet MS' >"& rsInfo("ItemDescription") &"</font></div></td> <td>" & _
                " <div align='center'><font face='Trebuchet MS' >"& rsInfo("QtyReceived") &"</font></div></td><td>" & _
                " <div align='center'><font face='Trebuchet MS' >"& rsInfo("QtyAccepted") &"</font></div></td><td>" & _
                " <div align='center'><font face='Trebuchet MS' >"& rsInfo("QtyRejected") &"</font></div></td></tr>"

			i = i + 1
			rsInfo.movenext
			Wend
			rsInfo.Close
			rsReqester.Close
			rsEmp.Close

    sBody = sBody & " </table> </td></tr>" & _
            " <tr> " & _
            " <td colspan='3' bgcolor=#DFF2FC><font face='Trebuchet MS'><b>Remarks : </b> "& rsGRN("Remarks") &"</font> </td>" & _
			"</tr><tr bgcolor=#108ed6>" &_
			"<td colspan='9' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail.</font></td>" &_
			"</tr>" &_
		    "</table>"

		'Response.write sBody
		eSubject = "GRN Details : " & GetGRNNo(GRNNum)
		eBody = sBody
		eBoolHtml=true
		call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

		End Function

%>

<script language="javascript">
	function redirect()
	{
		document.FinalForm.method="Post";
		document.FinalForm.action="Ack_GRN.asp"
		document.FinalForm.submit();
	}
</script>
<html>
<body onLoad="javascript:redirect();">
<form name="FinalForm">
<input type="hidden" name="hdGRN" value="<%=GRNNo%>">
</form>
</body>
</html>


<!--#include file="../includes/connection_close.asp"-->
