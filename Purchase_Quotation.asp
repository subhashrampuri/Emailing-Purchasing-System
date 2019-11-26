<%@LANGUAGE="VBSCRIPT"%>
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


<Script language=JavaScript src="../includes/javascript/validate.js" type=text/javascript></SCRIPT>
<%
	ReceivedEditAction="Add"
	EmployeeId=Session("Employee_Id")

	sql="select EmployeeId from tbl_WKI_Admin where EmployeeId='" & EmployeeId  & "'"
	call RunSQL(sql,rsAdmin)
	if not rsAdmin.eof then
		isAdmin=true
	else
		isAdmin=false
	end if
	rsAdmin.close()

	sql="select DeptName from VSPL_Managers A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsDepartment)
	if not rsDepartment.eof then
		isManager=true
		if rsDepartment(0)="HR" then
			isHRManager=true
		elseif rsDepartment(0)="ITIT" then
			isITITManager=true
		else
			isHRManager=false
			isITITManager=false
		end if
	else
		isManager=false
	end if
	rsDepartment.close()

	sql="select DeptName from VSPL_EmployeeHREntry A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsDepartment)
	if not rsDepartment.eof then
		if rsDepartment(0)="HR" then
			isHRMember=true
		elseif rsDepartment(0)="ITIT" then
			isITITMember=true
		else
			isHRMember=false
			isITITMember=false
		end if
	else
		isManager=false
	end if
	rsDepartment.close()

	PurRequisitionNo = Request.QueryString("PurRequisitionNo")

	sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& PurRequisitionNo &" "
	call RunSql(sql,rsRec)
	if rsRec.Eof = false then
		ReqNum = rsRec("RequisitionNum")
	end if
	rsRec.Close


	Dim CurPrj
	CurPrj = Request.Form("ItemDesc")

	Private Function fsSendMail_ApproverTeam()

	Dim sBody
	'-----Active Approver---------------
	sql="sp_PSystem_GetActiveApprover"
	call RunSql(sql,rsApprover)
	if rsApprover.eof then
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in Approvers Panel.</b></font></center>"
		Response.end
	end if
	ApproverId=rsApprover("EmployeeId")
	eToName=rsApprover("EmployeeName")
	eToEmail=rsApprover("EmployeeEmail")
	rsApprover.close()
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
		" </b><br>" & _
		" The synopses of the quotations entered are as follows: " & _
		"</font><br><br>" & _
		"<table width='100%' cellpadding='1' cellspacing='2'>" & _
		" <tr bgcolor=#108ed6> " & _
	    " <td colspan=10 align='center' height='25'><font face='Trebuchet MS' color='#ffffff'><b> Quotations Entered </b></font></td>" & _
		" </tr> <tr> <td colspan=10>&nbsp; </td> </tr> "

	  Dim i
	   Sql= "select distinct ItemDescription,ProjectId,RequisitionId from tbl_Psystem_Quotations where RequisitionId= "& PurRequisitionNo &" And isApproved = 0 and isPRCancelled = 0"
	   Call RunSql(sql,rsItems)

		while Not rsItems.EOF

			iDesc=rsItems("ItemDescription")
			PrjId = rsItems("ProjectId")
			sql = sql_GetProjectName(PrjID)
			call RunSql(sql,rsPrj)
			if not rsPrj.EOF then
				sPrjName = rsPrj("ProjectName")
				rsPrj.Close
			end if

    sBody = sBody & " <tr> <td align='center' colspan='10'>" & _
      	" <table width='100%' border='0' cellspacing='0' cellpadding='0'> " & _
        " <tr bgcolor=#108ed6> <td width='17%' nowrap> <div align='right'><font face='Trebuchet MS' color='#ffffff'><b>Requisition No : </b></font></div> </td>" & _
        " <td width='17%' nowrap><font color='#ffffff' face='Trebuchet MS'> " & GetPurchaseRequisitionNo(ReqNum) & "</font></td> " & _
        " <td width='17%' nowrap> <div align='right'><font face='Trebuchet MS' color='#ffffff'><b>Item Description :</b></font></div> </td>" & _
        " <td width='17%' nowrap><font face='Trebuchet MS' color='#ffffff'>" & iDesc & "</font></td><br> " & _
		" <td width='17%' nowrap> <div align='right'><font face='Trebuchet MS' color='#ffffff'><b>Project :</b></font></div> </td>" & _
        " <td width='15%' nowrap><font face='Trebuchet MS' color='#ffffff'>" & sPrjName & " <font></td> </tr> </table> </td> </tr> " & _
  		" <tr bgcolor=#108ed6> <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Sl. No.</b></font></td> " & _
		" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Supplier Name</b></font></td> " & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Unit Price</b></font></td>" & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Tax</b></font></td> " & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Tax Percent</b></font></td>" & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Quantity</b></font></td>" & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Warranty</b></font></td>" & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Delivery Time</b></font></td>" & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Payment Terms</b></font></td>" & _
    	" <td align='center'><font face='Trebuchet MS' color='#ffffff'><b>Remarks</b></font></td>" & _
  		" </tr> "

		sql= "Select * from tbl_Psystem_Quotations where ItemDescription = '" & Replace(Server.HTMLEncode(iDesc),"'","''") & "' and ProjectId = "& PrjId &" and RequisitionId = "& PurRequisitionNo &" And isApproved = 0 and isPRCancelled = 0 "
		Call RunSql(sql,rsInfo)
		i = 1
		while NOT rsInfo.EOF
			if rsInfo("Currency") = -1 then
				sCurr = "Rupee(s)"
			else
				sCurr = "Doller(s)"
			end if
			if rsInfo("isTaxIncludedOrExcluded") = -1 then
				sTax = "Exclusive"
			else
				sTax = "Inclusive"
			end if
			iCode = rsInfo("ItemCode")

  	sBody = sBody & "<tr bgcolor=#DFF2FC> " & _
    	" <td align='center'><font face='Trebuchet MS'> " & i & " </font></td> " & _
    	" <td align='left' style='word-break: break-all; width:200px;' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("SupplierName") & " </font></td> " & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("UnitPrice") & " " & sCurr & "</font> </td> " & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & sTax & "</font></td>" & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("TaxPercent") & " " & "%" & "</font> </td>" & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("Quantity") & "</font></td>" & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("Warranty") & "</font></td>" & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("DeliveryTime") & " </font></td>" & _
  	  	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("PaymentTerms") & "</font></td>" & _
    	" <td align='center' vAlign='top'> <font face='Trebuchet MS'>" & rsInfo("Remarks") & " </font></td>" & _
 		" </tr> "
			i = i + 1
			rsInfo.movenext
			Wend
			rsInfo.Close
  	sBody = sBody & "<tr> <td colspan=10>&nbsp; </td> </tr> "
			rsItems.movenext
			wend
			rsItems.close

	sBody = sBody &	"</tr><tr bgcolor=#108ed6>" &_
			"<td colspan='10' align='left'> <font face='Trebuchet MS' color='#ffffff'>This is an application automated e-mail. Please do not reply to this e-mail. </font></td>" &_
			"</tr>" &_
		    "</table>"


	'Response.write sBody
	eSubject = "Quotations for Approved Purchase Request : " & GetPurchaseRequisitionNo(ReqNum)
	eBody = sBody
	eBoolHtml=true
	Call SendEmail(eToName,eToEmail,eFromName,eFromEmail,eSubject,eBody,eCCName,eCCEmail,eBCCName,eBCCEmail,eAttachedFile, eBoolHTML)

	End Function
%>
<Script language="javascript">
	function Showhide()
	{
	var obj = document.getElementById("txp");
	if(document.strFormm.Tax.value == "Exclusive")
		obj.style.display = "";
		else
		obj.style.display = "none";
		document.strFormm.TaxPercent.value = "0";
	}

	function ItemInfo()
	{
		var ReqNo = (document.strFormm.hdReqNo.value);
		document.ItemList.method  ="post";
		document.ItemList.action="Purchase_Quotation.asp?PurRequisitionNo="+ReqNo;
		document.ItemList.submit();

	}
</script>
      <table width="100%" cellspacing="1" cellpadding="2" border="0">
        <tr height="25" class="blue" align="center">
          <td align="center" width="95%"> <font color=#ffffff><b>Quotation Entry Screen</b></font>
          </td>
          <td align="center" width="5%"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></td>
        </tr>
        <tr height="25" align="left">
          <td colspan="2" >
            <form name="ItemList">
              <table height="25" width="100%" border="0" cellpadding="0" cellspacing="2">
                <tr class="blue">
                  <%
				Dim sPrjName,sCurrency
				if trim(CurPrj) <> "" then

					sql = "Select * from tbl_PSystem_PurchaseRequestTransaction where (status = 1 or status = 2 or Status=3 ) and RequisitionId = " & PurRequisitionNo & " And ProjectId = " & Curprj & " "

					'sql = "Select ProjectId,RupeeOrDollar,PurchaseOrService,ApproxUnitCost,Purpose,RequiredDate,PossibleSource,SpecialInstruction from " &_
					' 	 "tbl_PSystem_PurchaseRequestTransaction where (status = 1 or status = 2) and RequisitionId = " & PurRequisitionNo & " And ProjectId = " & Curprj & " "

					Call RunSql(sql,rsApp)
					if not rsApp.eof  then
						sPrjID = rsApp("ProjectId")
						sql = sql_GetProjectName(sPrjID)
						call RunSql(sql,rsProject)
						if not rsProject.EOF then
							sPrjName = rsProject("ProjectName")
						end if
						if rsApp("RupeeOrDollar") = 0 then
							sCurrency = "Rupee(s)"
						else
							sCurrency = "Dollar(s)"
						end if
						if rsApp("PurchaseOrService") = 0 then
							sPur_Ser = "Purchase"
						else
							sPur_Ser = "Service"
						end if

						UnitCost=rsApp("ApproxUnitCost")
						purpose = rsApp("Purpose")
						ReqDate = SetDateFormat(rsApp("RequiredDate"))
						PossibleSource = rsApp("PossibleSource")
						SplInstruction = rsApp("SpecialInstruction")

					end if
					rsApp.close

					sql = "select sum(QuantityApproved) as QtyApproved from  tbl_Psystem_TransactionDetails where  RequisitionId = " & PurRequisitionNo & " and ProjectId = " & Curprj & " and isQuotationEntered = 0 and  (status = 2 or Status =1) "
					Call RunSql(sql,rsQty)

					if rsQty.EOF = false then
						QuantityApproved = rsQty("QtyApproved")
					rsQty.Close
					End if

				end if

				%>
                  <td align="center"><font color="#ffffff"><b>Item Description</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Project</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Purpose</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Quantity Approved</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Purchase/Service</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Requested Date</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Unit Price</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Possible Source</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Special Instructions</b></font></td>
                </tr>
                <tr height="25" bgcolor="<%=gsBGColorLight%>">
                  <td align="center" vAlign="top">
                    <Select class="formstylemed" name="ItemDesc" onChange="ItemInfo(this.value)"   onFocus="javascript:isToPropagate=true;">
                      <option value="0" Selected>Select Item Description</option>
                      <%
					sql = "Select distinct ItemDescription,ProjectId from tbl_Psystem_TransactionDetails where (status = 1 or status =2) and isquotationEntered = 0 and RequisitionId = " & PurRequisitionNo & " "
					Call RunSql(sql,rsItem)
				if not rsItem.eof then
					While not rsItem.eof
					if Cint(rsItem("ProjectId")) = Cint(CurPrj)	then %>
                      <option  Selected value="<%=rsItem("ProjectId")%>" ><%=rsItem("ItemDescription")%></option>
                      <%	else %>
                      <option value="<%=rsItem("ProjectId")%>"><%=rsItem("ItemDescription")%></option>
                      <%	end if	%>
                      <%
					rsItem.MoveNext
					Wend
					rsItem.Close
				else
					Call fsSendMail_ApproverTeam()
					Response.Redirect ("PurchaseApproveList.asp")
				end if
				%>
                    </select>
                  </td>
                  <td align="center" vAlign="top"><%=sPrjName%></td>
                  <td align="center" vAlign="top"><%=purpose%></td>
                  <td align="center" vAlign="top"><%=QuantityApproved%></td>
                  <td align="center" vAlign="top"><%=sPur_Ser%></td>
                  <td align="center" vAlign="top"><%=ReqDate%></td>
                  <td align="center" vAlign="top"><%=UnitCost & " " & sCurrency%></td>
                  <td align="center" vAlign="top"><%=PossibleSource%></td>
                  <td align="center" vAlign="top"><%=SplInstruction%></td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
        <tr valign="top">
          <td align="center" valign="top" colspan="2">
            <form name="strFormm">
              <table width="90%" align="center" id="QuotationList" valign="top">
                <tr align="center">
                  <td colspan="4"><b><%=ReceivedEditAction%> Quotation </b></td>
                </tr>
                <tr align="center">
                  <td colspan="4"><b><font color="red"><b>*</b></font><b>&nbsp;Fields
                    are mandatory </b> </b></td>
                </tr>

                <tr>
                  <td class="blue" align="right"><font color=#ffffff><b>Purchase
                    Requisition No&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <select class="formstylemed" name="RequisitionNo" onblur="javascript:validateRequisitionNo(this);" onfocus="javascript:isToPropagate=true;" >
                      <option value="<%=ReqNum%>" ><%=GetPurchaseRequisitionNo(ReqNum)%></option>
                    </select>
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Vendor/Supplier&nbsp;:<font color="red">*</font>
                    </b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <%
						sql = "Select SupplierName from tbl_Psystem_Supplier where isActive=1 order by SupplierName"
						Call RunSql(sql,rsSupplier)
					%>
                    <select class="formstylemed" name="Supplier"  onFocus="javascript:isToPropagate=true;">
                      <option value="0">Select Supplier</option>
                      <%	while NOT rsSupplier.EOF	%>
                      <option value="<%=rsSupplier("SupplierName")%>"><%=rsSupplier("SupplierName")%></option>
                      <%
					  	rsSupplier.movenext
						Wend
						rsSupplier.Close
					  %>
                    </select>
                  </td>
                </tr>
                <tr>
                  <td class="blue" align="right"><font color=#ffffff><b>Unit Price&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <input class="formstylemedium" type="text" size="25" maxlength="9" name="Price"  onFocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Currency&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <select class="formstylemed" name="Currency"  onFocus="javascript:isToPropagate=true;">
                      <option value="0">Select Currency</option>
                      <option value="Rupee(s)">Rupee(s)</option>
                      <option value="Doller(s)">Doller(s)</option>
                    </select>
                    <b></b> </td>
                </tr>
                <tr>
                  <td class="blue" align="right"><font color=#ffffff><b>Tax /
                    Excl Tax &nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <select class="formstylemed" name="Tax" onChange="Showhide()"  onFocus="javascript:isToPropagate=true;">
                      <option value="0" selected>Select Tax</option>
                      <option value="Inclusive">Inclusive</option>
                      <option value="Exclusive">Exclusive</option>
                    </select>
                    <font color=red> </font>
                    <div id="txp" name="txp" style="display: none;">
                      <input type="text" class=formstyleTooShort name="TaxPercent" value="0" maxlength="2" style="border: 1 solid" onBlur="javascript:validateTaxPercent(this);" >
                      in Percentage <font color=red> </font></div>
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Delivery
                    Time :<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <input class="formstylemedium" type="text" size="25" name="DeliveryTime"  maxlength="50" value=""  onFocus="javascript:isToPropagate=true;">
                  </td>
                </tr>
                <tr>
                  <td class="blue" align="right"><font color=#ffffff><b>Warranty&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <input class="formstylemedium" type="text" size="25" name="Warranty" maxlength="50"  onFocus="javascript:isToPropagate=true;">
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Quantity&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <input class="formstylemedium" type="text" size="25" name="Quantity" maxlength="4" READONLY value="<%=QuantityApproved%>"  onFocus="javascript:isToPropagate=true;">
                  </td>
                </tr>
                <tr>
                  <td class="blue" align="right"><font color=#ffffff><b>Payment
                    Terms&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <textarea class="formstylemedium" name="PaymentTerms" rows="4" cols="19"  onFocus="javascript:isToPropagate=true;"
						onKeyDown="textCounter(document.strFormm.PaymentTerms,document.strFormm.remLen3,500)"	onKeyUp="textCounter(document.strFormm.PaymentTerms,document.strFormm.remLen3,500)""></textarea>
                    <font color="red"> Max Chars (500)</font>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Remarks:&nbsp;</b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <textarea class="formstylemedium" name="Remarks" rows="4" cols="19"  onFocus="javascript:isToPropagate=true;"
						onKeyDown="textCounter(document.strFormm.Remarks,document.strFormm.remLen3,500)" onKeyUp="textCounter(document.strFormm.Remarks,document.strFormm.remLen3,500)""></textarea>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                    <font color="red">Max Chars (500)</font></td>
                </tr>
                <tr>
                  <td  align="right">&nbsp;</td>
                  <td >&nbsp;</td>
                  <td  align="right">&nbsp;</td>
                  <td >&nbsp;</td>
                </tr>
                <tr>
                  <td  align="center" colspan="4">
                    <input class="formbutton" type="button" name="AddButton" value="Add" style="border: 1 solid; width:50Px" onclick="javascript:addItemToQuotationList()">
                    &nbsp;
                    <input class="formbutton" type="button" name="ResetButton" value="Reset" style="border: 1 solid; width:50Px" onclick="javascript:resetItemDetails()">
                    &nbsp; </td>
                </tr>
              </table>
              <input type="hidden" name="EditAction" value="">
              <input type="hidden" name="hdReqNo" value="<%=PurRequisitionNo%>">
            </form>
          </td>
        </tr>
        <tr>
          <td colspan="2">
            <form name="ItemForm">
              <table id="QuotationRequest" width="100%" align="center" valign="top" cellspacing="2" cellpadding="2" border="0">
                <tr height="25" class="blue">
                  <td colspan="13" align="center"> <font color=#ffffff><b>Quotations
                    Filed</b></font></td>
                </tr>
                <tr height="25" class="blue">
                  <td align="center"> <font color=#ffffff><b>Sl. No.</b></font>
                  </td>
                  <td align="center"> <font color=#ffffff><b>Item Description</b></font>
                  <td align="center"> <font color=#ffffff><b>Requisition No</b></font>
                  </td>
                  <td align="center"> <font color=#ffffff><b>Supplier</b></font>
                  </td>
                  <td align="center"> <font color=#ffffff><b>Price</b></font>
                  </td>
                  <td align="center"> <font color=#ffffff><b>Currency</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Tax</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Tax in %</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Delivery Time</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Warranty</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Quantity</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Payment Terms</b></font></td>
                  <td align="center"> <font color=#ffffff><b>Remarks</b></font></td>
                </tr>
                <tr>
                  <td class="bluelight" colspan="13" align="center">
                    <input type="button" class="formbutton" value="Edit" style="border: 1 solid; width:50Px" onclick="javascript:editQuotationEntry();">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" class="formbutton" value="Delete" style="border: 1 solid; width:50Px" onclick="javascript:deleteItemFromQuotationEntry();">
                    &nbsp;&nbsp;&nbsp;&nbsp;
                    <input type="button" class="formbutton" value="Submit" style="border: 1 solid; width:50Px" onclick="javascript:validateOnSubmit();">
                  </td>
                </tr>
                <input name='CheckItem' type='checkbox' value="" style='visibility:hidden'>
                <input name='CheckItem' type='checkbox' value="" style='visibility:hidden'>
              </table>
            </form>
            <form name="FinalForm">
              <input type="hidden" name="ItemList" value="">
            </form>
          </td>
        </tr>
      </table>
	<script language="javascript">
		var isToPropagate=true;
		var index=2;
		var editIndex=-1;

		function validateItemDescription(ctrl)
		{
			if(isToPropagate==true)
			{
				if(ctrl.value=="0")
				{
					alert("Item Description is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}

		function validateRequisitionNo(ctrl)
		{
			if(isToPropagate==true)
			{
				if(ctrl.value=="0")
				{
					alert("Requisition No is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}

		function validateSupplier(ctrl)
		{
			if(isToPropagate==true)
			{
				if(ctrl.value=="0")
				{
					alert("Vendor/Supplier is a required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateTaxPercent(ctrl)
		{
			var regexp = new RegExp (/^[0-9]\d*$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Price enter Tax percent");
					isToPropagate=false;
					ctrl.value = "0";
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter a valid Tax Percent");
					isToPropagate=false;
					ctrl.value="0";
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;

		}
		function validatePriceRequired(ctrl)
		{
		//	var regexp = new RegExp (/^[1-9]\d*$/);
			var regexp = new RegExp (/^\d+(\.\d\d)?$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Price is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(ctrl.value < 1)
				{
					alert("Price cannot be zero");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter a valid Price Required.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}

		function validateCurrency(ctrl)
		{
			if(isToPropagate==true)
			{
				if(ctrl.value=="0")
				{
					alert("Currency is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateTax(ctrl)
		{
			if(isToPropagate==true)
			{
				if(ctrl.value=="0")
				{
					alert("Select Tax type");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}


		function validateDeliveryTime(ctrl)
		{
			var regexp = new RegExp (/^[1-9]\d*$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Delivery Time is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateWarranty(ctrl)
		{
			var regexp = new RegExp (/^[1-9]\d*$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Warranty is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}

		function validateQuantityRequired(ctrl)
		{
			var regexp = new RegExp (/^[1-9]\d*$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Quantity is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter a valid Quantity Required.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validatePaymentTerms(ctrl)
		{
			var regexp = new RegExp (/[0-9a-zA-Z]/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Payment Terms is required field");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter Alpha-Numeric characters only.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateRemarks(ctrl)
		{
			var regexp = new RegExp (/[0-9a-zA-Z]/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(ctrl.value!="")
			{
				if(isToPropagate==true)
				{
					if(!regexp.test(ctrl.value))
					{
						alert("Please enter Alpha-Numeric characters only.");
						isToPropagate=false;
						ctrl.focus();
						return false;
					}
					else
						isToPropagate=true;
				}
			}
			return true;
		}

		function validateOnAdd()
		{
			var referenceItemDesc=document.getElementById("ItemDesc");
			var referenceReqNo=document.getElementById("RequisitionNo");
			var referenceSupplier=document.getElementById("Supplier");
			var referencePrice=document.getElementById("Price");
			var referenceCurrency=document.getElementById("Currency");
			var referenceTax=document.getElementById("tax");
			var referenceDeliveryTime=document.getElementById("DeliveryTime");
			var referencePaymentTerms=document.getElementById("PaymentTerms");
			var referenceWarranty=document.getElementById("Warranty");
			var referenceQuantity=document.getElementById("Quantity");
			var referenceRemarks=document.getElementById("Remarks");

			var returnValueItemDesc=validateItemDescription(referenceItemDesc);
			var returnValueReqNo=validateRequisitionNo(referenceReqNo);
			var returnValueSupplier=validateSupplier(referenceSupplier);
			var returnValuePrice=validatePriceRequired(referencePrice);
			var returnValueCurrency=validateCurrency(referenceCurrency);
			var returnTax=validateTax(referenceTax);
			var returnValueDeliveryTime=validateDeliveryTime(referenceDeliveryTime);
			var returnValuePaymentTerms=validatePaymentTerms(referencePaymentTerms);
			var returnValueWarranty=validateWarranty(referenceWarranty);
			var returnValueQuantity=validateQuantityRequired(referenceQuantity);
			var returnValueRemarks=validateRemarks(referenceRemarks);

			if(returnValueItemDesc==true && returnValueReqNo==true && returnValueSupplier==true && returnValuePrice==true && returnValueCurrency==true && returnTax==true && returnValueDeliveryTime==true && returnValuePaymentTerms==true && returnValueWarranty==true && returnValueQuantity==true && returnValueRemarks==true)
			{
				return true;
			}
			else
				return false;
		}
		function validateOnSubmit()
		{
			if(document.ItemForm.CheckItem.length==2)
			{
				alert("Quotation Request is empty");
				return false;
			}
			else if(document.ItemForm.CheckItem.length==3)
			{
				alert("Minimum 2 Quotations are mandatory")
				return false;
			}
			else
			{
				var tableReference=document.getElementById('QuotationRequest');
				var rowCount=tableReference.rows.length;
				--rowCount;
				var arrItems=new Array(rowCount-2);

				for(var i=2;i<rowCount;i++)
				{
					arrItems[i-2]=new Array();
					for(var j=0;j<=12;j++)
					{
						if (j==0)
							arrItems[i-2][j] = document.ItemList.ItemDesc.value
						else
						//	arrItems[i-2][j]=tableReference.rows[i].cells[j].innerText;
						{
							var temp=tableReference.rows[i].cells[j].innerText;
							while(temp.search(",")!=-1)
								temp=temp.replace(",","&#44;");
							arrItems[i-2][j]=temp;
						}


					}

				}
			//	alert(arrItems);
				document.FinalForm.ItemList.value=arrItems;
				document.FinalForm.method="post";
				document.FinalForm.action="Submit_Quotation.asp";
				document.FinalForm.submit();
				return true;
			}
		}

		function addItemToQuotationList_1()
		{
			for (var i=0; i < document.ItemList.ItemDesc.length;i++)
			{
			if (document.ItemList.ItemDesc.options[i].selected == true)
				{
					alert(document.ItemList.ItemDesc.options[i].text)
				}
			}
			alert(document.strFormm.RequisitionNo.value);
			alert(document.strFormm.Supplier.value);
			alert(document.strFormm.Price.value);
			alert(document.strFormm.Currency.value);
			for (var j=0; j < document.strFormm.Tax.length;j++)
			{
			if (document.strFormm.Tax.options[j].selected == true)
				{
					alert(document.strFormm.Tax.options[j].text)
				}
			}
			if (document.strFormm.Tax.value == 1)
			{
			alert(document.strFormm.TaxPercent.value);
			}
			else
			{
			document.strFormm.TaxPercent.value = 0;
			alert(document.strFormm.TaxPercent.value);
			}
			alert(document.strFormm.DeliveryTime.value);
			alert(document.strFormm.Warranty.value);
			alert(document.strFormm.Quantity.value);
			alert(document.strFormm.PaymentTerms.value);
			alert(document.strFormm.Remarks.value);
		}

		function addItemToQuotationList()
		{
			if(document.getElementById("AddButton").value=="Add")
			{
				if(validateOnAdd()==true)
				{

					document.getElementById('QuotationRequest').insertRow(index);
					for(var i=0;i<13;i++)
					{
						document.getElementById('QuotationRequest').rows[index].insertCell();
					}
					document.getElementById('QuotationRequest').rows[index].cells[0].innerHTML="<input name='CheckItem' type='checkbox' value='" + index + "'>"+(document.getElementById('QuotationRequest').rows[index].rowIndex-1);
					document.getElementById('QuotationRequest').rows[index].cells[0].align="center";

					document.getElementById('QuotationRequest').rows[index].cells[0].className="bluelight";
					for (var i=0; i < document.ItemList.ItemDesc.length;i++)
						{
							if (document.ItemList.ItemDesc.options[i].selected == true)
								{
									var sDesc = (document.ItemList.ItemDesc.options[i].text);
								}
						}
					document.getElementById('QuotationRequest').rows[index].cells[1].innerText=sDesc;
					document.getElementById('QuotationRequest').rows[index].cells[1].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[1].className="bluelight";

					//document.getElementById('QuotationRequest').rows[index].cells[2].innerText=document.strFormm.RequisitionNo.value;
					document.getElementById('QuotationRequest').rows[index].cells[2].innerText='<%=GetPurchaseRequisitionNo(ReqNum)%>';
					document.getElementById('QuotationRequest').rows[index].cells[2].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[2].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[3].innerText=document.strFormm.Supplier.value;
					document.getElementById('QuotationRequest').rows[index].cells[3].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[3].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[4].innerText=document.strFormm.Price.value;
					document.getElementById('QuotationRequest').rows[index].cells[4].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[4].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[5].innerText=document.strFormm.Currency.value;
					document.getElementById('QuotationRequest').rows[index].cells[5].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[5].className="bluelight";
					for (var j=0; j < document.strFormm.Tax.length;j++)
						{
							if (document.strFormm.Tax.options[j].selected == true)
								{
									var sTax = (document.strFormm.Tax.options[j].text);
								}
						}
					document.getElementById('QuotationRequest').rows[index].cells[6].innerText=sTax;
					document.getElementById('QuotationRequest').rows[index].cells[6].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[6].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[7].innerText=(document.strFormm.TaxPercent.value);
					document.getElementById('QuotationRequest').rows[index].cells[7].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[7].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[8].innerText=document.strFormm.DeliveryTime.value;
					document.getElementById('QuotationRequest').rows[index].cells[8].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[8].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[9].innerText=document.strFormm.Warranty.value;
					document.getElementById('QuotationRequest').rows[index].cells[9].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[9].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[10].innerText=document.strFormm.Quantity.value;
					document.getElementById('QuotationRequest').rows[index].cells[10].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[10].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[11].innerText=document.strFormm.PaymentTerms.value;
					document.getElementById('QuotationRequest').rows[index].cells[11].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[11].className="bluelight";
					document.getElementById('QuotationRequest').rows[index].cells[12].innerText=(document.strFormm.Remarks.value=="")?"NA":document.strFormm.Remarks.value;
					document.getElementById('QuotationRequest').rows[index].cells[12].align="center";
					document.getElementById('QuotationRequest').rows[index].cells[12].className="bluelight";

					index++;
					resetItemDetails();
				}
				else
					return false;
			}
			else if(validateOnAdd()==true)
			{
					for (var i=0; i < document.ItemList.ItemDesc.length;i++)
					{
					if (document.ItemList.ItemDesc.options[i].selected == true)
						{
							var sDesc = (document.ItemList.ItemDesc.options[i].text);
						}
					}
					document.getElementById('QuotationRequest').rows[editIndex].cells[1].innerText=sDesc;
					//document.getElementById('QuotationRequest').rows[editIndex].cells[2].innerText=document.strFormm.RequisitionNo.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[2].innerText='<%=GetPurchaseRequisitionNo(ReqNum)%>';
					document.getElementById('QuotationRequest').rows[editIndex].cells[3].innerText=document.strFormm.Supplier.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[4].innerText=document.strFormm.Price.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[5].innerText=document.strFormm.Currency.value;
					for (var j=0; j < document.strFormm.Tax.length;j++)
						{
							if (document.strFormm.Tax.options[j].selected == true)
								{
									var sTax = (document.strFormm.Tax.options[j].text);

								}
						}
					document.getElementById('QuotationRequest').rows[editIndex].cells[6].innerText=sTax;
					document.getElementById('QuotationRequest').rows[editIndex].cells[7].innerText=(document.strFormm.TaxPercent.value);
					document.getElementById('QuotationRequest').rows[editIndex].cells[8].innerText=document.strFormm.DeliveryTime.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[9].innerText=document.strFormm.Warranty.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[10].innerText=document.strFormm.Quantity.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[11].innerText=document.strFormm.PaymentTerms.value;
					document.getElementById('QuotationRequest').rows[editIndex].cells[12].innerText=(document.strFormm.Remarks.value=="")?"NA":document.strFormm.Remarks.value;

					editIndex=-1;
					document.getElementById('AddButton').value="Add";
					document.getElementById('ResetButton').disabled=false;
					resetItemDetails();
					return true;

			}
		}
		function resetItemDetails()
		{
			//var x=document.getElementById("Supplier")
			  // x.remove(x.selectedIndex)
			document.strFormm.Supplier.value="0";
			document.strFormm.Price.value="";
			document.strFormm.Currency.value="0";
		//	document.strFormm.Tax.value="Inclusive";
			document.strFormm.Tax.value="0";
			Showhide();
			document.strFormm.TaxPercent.value="0";
			document.strFormm.DeliveryTime.value="";
			document.strFormm.PaymentTerms.value="";
			document.strFormm.Warranty.value="";
			//document.strFormm.Quantity.value="";
			document.strFormm.Remarks.value="";
		}
		function deleteItemFromQuotationEntry()
		{
			var selectedItemFlag=false;
			var selectedItemIndex;
			if(document.ItemForm.CheckItem.length==2)
			{
				alert("Quotation entry is empty");
				return false;
			}
			else
			{
				var checkedCount=0;
				for(var i=0;i<document.ItemForm.CheckItem.length-2;i++)
					if(document.ItemForm.CheckItem[i].checked==true)
					{
						if(selectedItemFlag==false)
						{
							selectedItemIndex=i;
							selectedItemFlag=true;
						}
						checkedCount++;
					}
				if(checkedCount==0)
				{
					alert("Please select an item to edit.");
					return false;
				}
				else
				{
					var tableReference=document.getElementById('QuotationRequest');
					var rowCount;
					var reOrder=false;
					for(var i=document.ItemForm.CheckItem.length-3;i>=0;i--)
					{
						if(document.ItemForm.CheckItem[i].checked==true)
						{
							tableReference.deleteRow(document.ItemForm.CheckItem[i].value);
							--index;
							reOrder=true;
						}
					}
					if(reOrder==true)
					{
						rowCount=tableReference.rows.length;
						--rowCount;
						for(var i=2;i<rowCount;i++)
							document.getElementById('QuotationRequest').rows[i].cells[0].innerHTML="<input name='CheckItem' type='checkbox' value='" + i + "'></input>&nbsp;"+(document.getElementById('QuotationRequest').rows[i].rowIndex-1);
					}
					return true;
				}
			}
		}
		function editQuotationEntry()
		{
			var selectedItemFlag=false;
			var selectedItemIndex;
			if(document.ItemForm.CheckItem.length==2)
			{
				alert("Quotation entry is empty");
				return false;
			}
			else
			{
				var checkedCount=0;
				for(var i=0;i<document.ItemForm.CheckItem.length-2;i++)
					if(document.ItemForm.CheckItem[i].checked==true)
					{
						if(selectedItemFlag==false)
						{
							selectedItemIndex=i;
							selectedItemFlag=true;
						}
						checkedCount++;
					}
				if(checkedCount==0)
				{
					alert("Please select an item to edit.");
					return false;
				}
				else if(checkedCount>1)
				{
					alert("Please select only one item to edit.");
					return false;
				}
				else
				{
					var referenceTableRow=document.getElementById('QuotationRequest').rows[document.ItemForm.CheckItem[selectedItemIndex].value];
					document.strFormm.Supplier.value=referenceTableRow.cells[3].innerText;
					document.strFormm.Price.value=referenceTableRow.cells[4].innerText;
					document.strFormm.Currency.value=referenceTableRow.cells[5].innerText;
					document.strFormm.Tax.value=referenceTableRow.cells[6].innerText;
						if (referenceTableRow.cells[6].innerText=="Exclusive")
						{
						//	alert(referenceTableRow.cells[7].innerText);
							Showhide();
							document.strFormm.TaxPercent.value=referenceTableRow.cells[7].innerText;

						}
					document.strFormm.DeliveryTime.value=referenceTableRow.cells[8].innerText;
					document.strFormm.Warranty.value=referenceTableRow.cells[9].innerText;
					document.strFormm.Quantity.value=referenceTableRow.cells[10].innerText;
					document.strFormm.PaymentTerms.value=referenceTableRow.cells[11].innerText;

					if (referenceTableRow.cells[12].innerText == "NA")
						document.strFormm.Remarks.value = "";
					else
					{
						document.strFormm.Remarks.value=referenceTableRow.cells[12].innerText;
					}

					editIndex=document.ItemForm.CheckItem[selectedItemIndex].value;
					document.ItemForm.CheckItem[selectedItemIndex].checked=false;
					document.getElementById('AddButton').value="Save"
					document.getElementById('ResetButton').disabled=true;
					return true;
				}
				return true;
			}
		}
		function callOpenCalendar(ctrl)
		{
			if(ctrl.value=="")
				openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
		}
		function ChangeToMMDDYYYY(strDate)
		{
			var datarr;
			var strmon;
			datarr=strDate.split("-")
			switch(datarr[1])
			{
				case 'Jan':
					strmon="01";
					break;
				case 'Feb':
					strmon="02";
					break;
				case 'Mar':
					strmon="03";
					break;
				case 'Apr':
					strmon="04";
					break;
				case 'May':
					strmon="05";
					break;
				case 'Jun':
					strmon="06";
					break;
				case 'Jul':
					strmon="07";
					break;
				case 'Aug':
					strmon="08";
					break;
				case 'Sep':
					strmon="09";
					break;
				case 'Oct':
					strmon="10";
					break;
				case 'Nov':
					strmon="11";
					break;
				case 'Dec':
					strmon="12";
					break;
			}
			if(datarr[0]=="1"||datarr[0]=="2"||datarr[0]=="3"||datarr[0]=="4"||datarr[0]=="5"||datarr[0]=="6"||datarr[0]=="7"||datarr[0]=="8"||datarr[0]=="9")
				datarr[0]="0"+datarr[0];

			return (strmon+'/'+datarr[0]+'/'+datarr[2]);
		}
		function ChangeToVelankaniFormatDate(strDate)
		{
			var datarr;
			var strmon;
			datarr=strDate.split("/")
			switch(datarr[0])
			{
				case "1":
					strmon="Jan";
					break;
				case "2":
					strmon="Feb";
					break;
				case "3":
					strmon="Mar";
					break;
				case "4":
					strmon="Apr";
					break;
				case "5":
					strmon="May";
					break;
				case "6":
					strmon="Jun";
					break;
				case "7":
					strmon="Jul";
					break;
				case "8":
					strmon="Aug";
					break;
				case "9":
					strmon="Sep";
					break;
				case "10":
					strmon="Oct";
					break;
				case "11":
					strmon="Nov";
					break;
				case "12":
					strmon="Dec";
					break;
			}
			return (datarr[1] + "-" + strmon + "-" + datarr[2]);
		}
	</script>
<SCRIPT LANGUAGE="JavaScript">

<!-- Web Site:  The JavaScript Source -->
<!-- Use one function for multiple text areas on a page -->
<!-- Limit the number of characters per textarea -->
<!-- Begin
function textCounter(field,cntfield,maxlimit) {
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);

// otherwise, update 'characters left' counter
else
cntfield.value = maxlimit - field.value.length;
}
//  End -->
</script>

<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->