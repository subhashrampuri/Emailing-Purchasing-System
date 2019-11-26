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

	Dim iReqId,sSupplier
	iReqId = Request.form("ItemDesc")
	'Response.write iReqId
	sSupplier = Request.Form("Suppliers")
%>
<Script language=JavaScript src="../includes/javascript/validate.js" type=text/javascript></SCRIPT>
<script language="javascript" type="text/javascript">
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

	function validateRequiredDate(ctrl)
	{
		var isToPropagate=true;
		var regexp=new RegExp(/(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		if(isToPropagate==true)
		{
			if(ctrl.value=="")
			{
				alert("Required Date is required field");
				isToPropagate=false;
			//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
				return false;
			}
			else if(!regexp.test(ChangeToMMDDYYYY(ctrl.value)))
			{
				alert("Please enter a valid Required Date.");
				isToPropagate=false;
			//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
				return false;
			}
			else if (isGreaterDate(ChangeToMMDDYYYY(document.strFormm.RequiredDate.value),ChangeToMMDDYYYY('<%=SetDateFormat(Formatdatetime(now(),2))%>')))
			{
				alert ("Required date should be Greater than or equal to Current Date")
				isToPropagate=false;
				ctrl.focus();
				return false;
			}

			else
				isToPropagate=true;
		}
			return true;
	}
	function validateItemDescription(ctrl)
	{
		if(ctrl.value=="0")
		{
			alert("Select Purchase order");
			ctrl.focus();
			return false;
		}
		return true;
	}
</script>

<Script language="Javascript">
	function ItemInfo(PrjId)
	{
	//	alert(PrjId);
		document.PurchaseOrder.method="post";
		document.PurchaseOrder.action="PurchaseOrder.asp";
		document.PurchaseOrder.submit();
	}
</Script>
<script language="Javascript">
	function Validator(frm)
	{
		var str,s,i
    	formElements=["RequiredDate","PaymentTerms"];
     	for(i=0;i<1;i++)
    	{
	      if(frm.elements[formElements[i]].value.length !=0)
    	  {
        	 str=frm.elements[formElements[i]].value
	         s = str.replace(/^(\s)*/, '');
	         s = s.replace(/(\s)*$/, '');
	         frm.elements[formElements[i]].value=s
    	  }
	    }

		//alert(document.PurchaseOrder.Itemdesc.value);
		//if(validateItemDescription(document.PurchaseOrder.Itemdesc)==false)
		//	return false;
		if (document.PurchaseOrder.ItemDesc.value=="0")
		{
			alert("Please select approved purchase request");
			document.PurchaseOrder.ItemDesc.focus();
			return false;
		}
		if (document.PurchaseOrder.Suppliers.value=="0")
		{
			alert("Please select supplier");
			document.PurchaseOrder.Suppliers.focus();
			return false;
		}
		if (frm.RequiredDate.value == "")
			{
			alert("Date is Required field");
			frm.RequiredDate.focus();
			return false;
			}
		if (frm.PaymentTerms.value == "")
			{
			alert("PaymentTerms are Required field");
			frm.PaymentTerms.focus();
			return false;
			}
		if (isGreaterDate(ChangeToMMDDYYYY(document.strFormm.RequiredDate.value),ChangeToMMDDYYYY('<%=SetDateFormat(Formatdatetime(now(),2))%>')))
			{
				alert ("Required date should be Greater than or equal to Current Date")
				frm.RequiredDate.focus();
				return false;
			}
		return true;
	}

</script>


      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr class="blue" align="center">
          <td align="center" colspan="2" width="95%"><font color=#ffffff><b>Purchase Order
            Screen</b></font> </td>
          <td align="center" width="5%"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></td>
        </tr>
        <tr height="25" align="left">
          <td colspan="3" >
            <form name="PurchaseOrder">
              <table width="100%" align="left" cellspacing="2" cellpadding="1">
                <tr class="blue">
                  <td align="left"><font color="#ffffff">&nbsp;<b>Approved Purchase
                    Requests</b></font></td>
                </tr>
                <tr bgcolor="<%=gsBGColorLight%>">
                  <td>
                    <Select class="formstylemed" name="ItemDesc" onChange="ItemInfo(this.value)" >
                      <option Selected value="0">Select Approved Request</option>
                      <%
						sql = "Select distinct RequisitionId from tbl_PSystem_Quotations where isApproved = 1 and isPROnHold = 0 and isPRCancelled = 0 "
						Call RunSql(sql,rsList)
						While Not rsList.EOF
						ReqID = rsList("RequisitionId")

						sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& ReqID &" "
						call RunSql(sql,rsRec)
						if rsRec.Eof = false then
							ReqNum = rsRec("RequisitionNum")
						end if
						rsRec.Close
						
						if cInt(rsList("RequisitionId")) = cInt(iReqId) then
					 %>
                      <option Selected value="<%=rsList("RequisitionId")%>"><%=GetPurchaseRequisitionNo(ReqNum)%></option>
                      <% else %>
                      <option value="<%=rsList("RequisitionId")%>"><%=GetPurchaseRequisitionNo(ReqNum)%></option>
                      <% end if %>
                      <%
						rsList.movenext
						Wend
						rsList.close
					  %>
                    </select>
                  </td>
                </tr>
                <tr height="25%">
                  <td class="blue"><font color="#ffffff">&nbsp;<b>Load Suppliers</b></font></td>
                </tr>
                <tr>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <select class="formstylemed" name="Suppliers" onChange="ItemInfo(this.value)">
                      <option selected value="0">Select Supplier</option>
                      <%
					if iReqId <> "" then
						sql= "Select distinct SupplierName from tbl_Psystem_Quotations where isApproved = 1 and RequisitionId = "& iReqId &" "
						call RunSql(sql,rsSup)
						While Not RsSup.Eof
						if cStr(rsSup("SupplierName")) = cStr(sSupplier) then %>
                      <option selected value="<%=rsSup("SupplierName")%>"><%=rsSup("SupplierName")%></option>
                      <% else %>
                      <option value="<%=rsSup("SupplierName")%>"><%=rsSup("SupplierName")%></option>
                      <% end if %>
                      <%
						rsSup.movenext
						Wend
						rsSup.close
					end if
					%>
                    </select>
                  </td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
        <tr>
          <td align="center" valign="top" colspan="3">
            <form name="Display">
              <table width="100%" cellspacing="2" callpadding="0" align="center">
                <tr class="blue">
                  <td align="center" ><b><font color="#ffffff">Sl.No</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Item Description</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Project</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Supplier Info</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Unit Price</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Tax Percent</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Quantity</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Warranty</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Delivery Time</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Payment Terms</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Remarks</font></b></td>
                  <td align="center" ><b><font color="#ffffff">Amount</font></b></td>
                </tr>
                <%
				Dim i
				if sSupplier <> "" then
					sql = "Select * from tbl_PSystem_Quotations where RequisitionId =  " & iReqID &" and SupplierName = '" & sSupplier & "' and isApproved =1"
					call RunSql(sql,rsItem)
					i = 1
					GTotal = 0
					while Not rsItem.EOF
						sPrjID = rsItem("ProjectId")
						sql = sql_GetProjectName(sPrjID)
						call RunSql(sql,rsProject)
						if not rsProject.EOF then
							sPrjName = rsProject("ProjectName")
						end if
						SupName = rsItem("SupplierName")
						sql= "Select SupplierAddress from tbl_Psystem_Supplier where SupplierName = '" & SupName & "' "
						Call RunSql(sql,rsSup)

						if Not rsSup.Eof then
							SupplierAddr = rsSup("SupplierAddress")
						end if
						rsSup.Close

						if rsItem("Currency") = -1 then
							sCurr = "Rs."
						else
							sCurr = "$"
						end if

						Qty = rsItem("Quantity")
						Price = rsItem("UnitPrice")
						TaxPercent = rsItem("TaxPercent")
						Tax = (cDbl(TaxPercent) / 100)
						Amount = (cInt(Qty) * cDbl(Price))

						if rsItem("isTaxIncludedOrExcluded") = -1 then
							Total = ((Amount) + ((Amount)* cDbl(Tax)))
						else
							Total = Amount
						end if

				%>
                <tr bgcolor="<%=gsBGColorLight%>">
                  <td align="center" vAlign="top"><%=i%></td>
                  <td align="center" vAlign="top" nowrap><%=rsItem("ItemDescription")%></td>
                  <td align="center" vAlign="top" nowrap><%=sPrjName%></td>
                  <td align="center" vAlign="top"><p style="word-break: break-all; width:200px;"><%=rsItem("SupplierName") & "<br>" & SupplierAddr%></p></td>
                   <td align="center" vAlign="top" nowrap><%=sCurr & " " & rsItem("UnitPrice")%></td>
                  <td align="center" vAlign="top"><%=rsItem("TaxPercent")%> %</td>
                  <td align="center" vAlign="top"><%=rsItem("Quantity")%></td>
                  <td align="center" vAlign="top"><%=rsItem("Warranty")%></td>
                  <td align="center" vAlign="top"><%=rsItem("DeliveryTime")%></td>
                  <td align="center" vAlign="top"><%=rsItem("PaymentTerms")%></td>
                  <td align="center" vAlign="top"><%=rsItem("Remarks")%></td>
                  <td align="center" vAlign="top" nowrap><%=sCurr & " " & FormatNumber(Total,2)%></td>
                </tr>
                <%
					i = i + 1
					GTotal = cDbl(GTotal) + cDbl(FormatNumber(Total,2))
					rsItem.movenext
					Wend
					rsItem.close
					end if
				%>
                <tr >
                  <td align="center" NOWRAP colspan="10">&nbsp;</td>
                  <td bgcolor="<%=gsBGColorLight%>" align="center" NOWRAP><b>Grand
                    Total</b></td>
                  <td bgcolor="<%=gsBGColorLight%>" align="Right" nowrap><b><%=sCurr & " " & FormatNumber(GTotal,2)%></b></td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
        <tr>
          <td colspan="3">
            <form name="strFormm" method ="Post" action="Submit_PurchaseOrder.asp" onSubmit='return Validator(this)'>
              <table cellpadding="2" cellspacing="2" align="center" width="45%">
                <tr >
                  <td colspan="2" align="center"><font color="red"><b>*</b></font><b>&nbsp;Fields
                    are mandatory</b></td>
                </tr>
                <tr class="blue">
                  <td colspan="2" align="center"><font color="#ffffff"><b>Terms and Conditions</b></font></td>
                </tr>
                <tr>
                  <td height="43" class="blue" align="right"><font color="#ffffff"><b>Date
                    of Delivery :<font color="red">*</font></b></font></td>
                  <td height="43" bgcolor="<%=gsBGColorLight%>">
                    <input class="formstylemedium" type="text" name="RequiredDate" size="11" value="" onfocus="javascript:isToPropagate=true;" Readonly>
                    &nbsp; &nbsp;<a name="CalanderLink" onClick="openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300)"><img border="0" src="/gif/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" width="16" height="15"></a>&nbsp;<b></b>&nbsp;
                  </td>
                </tr>
                <tr>
                  <td class="blue" align="right"><font color="#ffffff"><b>Payment
                    Terms :<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <textarea class="formstylemedium" name="PaymentTerms" rows="4" cols="20"
					onKeyDown="textCounter(document.strFormm.PaymentTerms,document.strFormm.remLen3,500)"  onKeyUp="textCounter(document.strFormm.PaymentTerms,document.strFormm.remLen3,500)"></textarea>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                    <font color="red">Max Chars (500)</font></td>
                </tr>
                <tr>
                  <td class="blue" align="right"><font color="#ffffff"><b>Others
                    : </b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <textarea class="formstylemedium" name="Others" rows="4" cols="20"
					onKeyDown="textCounter(document.strFormm.Others,document.strFormm.remLen3,500)"  onKeyUp="textCounter(document.strFormm.Others,document.strFormm.remLen3,500)"></textarea>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                    <font color="red">Max Chars (500)</font></td>
                </tr>
                <tr>
                  <td class="blue">&nbsp;</td>
                  <td bgcolor="<%=gsBGColorLight%>">
                    <p style="margin-left:50">
                      <input class="formbutton" type="Submit" name="AddButton" value="Submit" style="border: 1 solid; width:50px;" >
                      &nbsp;
                      <input class="formbutton" type="Reset" name="ResetButton" value="Reset" style="border: 1 solid; width:50Px;" >
                      <input type="hidden" name="hdReqId" value="<%=iReqId%>">
					  <input type="hidden" name="hdSupplier" value="<%=sSupplier%>">
					  <input type="hidden" name="hdSupAddr" value = "<%=SupplierAddr%>">
                      <input type="hidden" name="hdGTotal" value="<%=GTotal%>">
					 </p>
                  </td>
                </tr>
                <tr>
                  <td >&nbsp;</td>
                  <td align="right">&nbsp;</td>
                </tr>
              </table>
            </form>
          </td>
        </tr>
      </table>
<p align="center">
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