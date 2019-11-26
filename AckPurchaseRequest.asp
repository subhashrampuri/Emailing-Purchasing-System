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
	RequisitionId = Request.Form("hdReqId")

	sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& RequisitionId &" "
	call RunSql(sql,rsRec)
	if rsRec.Eof = false then
		ReqNum = rsRec("RequisitionNum")
	end if
	rsRec.Close

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
%>

<script language="JavaScript" type="text/javascript">
	function CallReqID()
	{
		myPopup = window.open('PurchaseRequest_Print.asp?lID=<%=RequisitionId%>&sReq=<%=GetPurchaseRequisitionNo(ReqNum)%>&sDept=<%=EmployeeDepartmentName%>&sEmpName=<%=EmployeeName%>&sEmpID=<%=EmployeeId%>',42,'toolbar=no,width=800,height=600,scrollbars=yes,left=100,top=150,resizable=yes');
		if (!myPopup.opener)
	    myPopup.opener = self;
	}
</script>
      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr height="25" class="blue" align="center">
          <td align="center"> <font color=#ffffff><b>Manager--Purchase Request</b></font>
          </td>
        </tr>
        <tr height="25"  align="Center">
          <td align="Right"><a href="javascript:CallReqID();"><img src="images/printer.gif" width="20" height="21" border="0"></a></td>
        </tr>
        <tr>
          <td>
            <table id="PurchaseRequest" width="98%" align="center" valign="top" cellspacing="2" cellpadding="2" border="0">
              <tr height="25" class="blue">
                <td colspan="10" align="center"> <font color=#ffffff><b>Purchase
                  Request: Acknowledgement</b></font> </td>
              </tr>
              <tr height="25" bgcolor="<%=gsBGColorLight%>">
                <td colspan="5">&nbsp;Employee: <font color='red'><%=EmployeeName%>
                  ( <%=EmployeeId%> )</font></td>
                <td colspan="5">&nbsp;Department: <font color='red'><%=EmployeeDepartmentName%></font></td>
              </tr>
              <tr height="25" bgcolor="<%=gsBGColorLight%>">
                <td colspan="5">&nbsp;Purchase Requisition No.: <font color='red'><%=GetPurchaseRequisitionNo(ReqNum)%></font></td>
                <td colspan="5">&nbsp;Requisition Date: <font color='red'><%=SetDateFormat(Date())%></font></td>
              </tr>
              <tr height="25" class="blue">
                <td align="center"> <font color=#ffffff><b>Sl. No.</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Item Description</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Project</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Purpose</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Quantity Required</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Request Type</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Required Date</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Approx Unit Cost</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Possible Source</b></font>
                </td>
                <td align="center"> <font color=#ffffff><b>Special Instructions</b></font>
                </td>
              </tr>
              <%
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
				%>
              <tr height="25" bgcolor="<%=lclstr_bgColor%>">
                <td align="center">
                  <%
							Response.write counter
							counter=counter+1
						%>
                </td>
                <td>
                  <% Response.write rsItems("ItemDescription") %>
                </td>
                <td>
                  <% Response.write rsItems("Project") %>
                </td>
                <td>
                  <% Response.write rsItems("Purpose") %>
                </td>
                <td align="center">
                  <% Response.write rsItems("QuantityRequested") %>
                </td>
                <td>
                  <% Response.write rsItems("ServiceType") %>
                </td>
                <td>
                  <% Response.write rsItems("RequiredDate") %>
                </td>
                <td>
                  <%
							if rsItems("ApproxUnitCost")= "" then
								Response.write "-"
							else
								Response.write rsItems("ApproxUnitCost") & " " & rsItems("Currency")
							end if
						%>
                </td>
                <td>
                  <% Response.write rsItems("PossibleSource") %>
                </td>
                <td>
                  <% Response.write rsItems("SpecialInstruction") %>
                </td>
              </tr>
              <%
					rsItems.movenext
					wend
					rsItems.close()
				%>
            </table>
          </td>
        </tr>
      </table>


<p align="center">
<a href="PurchaseRequest.asp" style="text-decoration:none"><b>Back</b></a> <br><br> <a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->