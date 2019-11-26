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

	Dim iReqId
	iReqId = Request.form("ItemDesc")
	'Response.write iReqId
%>




      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr class="blue" align="center">
          <td align="center" width="95%"> <font color=#ffffff><b>Released Purchase Orders
            </b></font></td>
          <td align="right" width="5%"><p style="margin-right:10"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></p></td>
        </tr>
        <tr >
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr align="left">
          <td colspan="2" > </td>
        </tr>
        <tr>
          <td align="center" >&nbsp; </td>
        </tr>
        <tr>
          <td colspan="2">
            <form name="strFormm">
              <table cellspaning="2" cellpadding="2" border="0" width="50%" align="center">
                <tr class="blue" align="center">
                  <td align="Center"><font color="#ffffff">Sl.No</font></td>
                  <td align="center"><font color="#ffffff">Purchase Order No</font></td>
                  <td align="center"><font color="#ffffff">Status</font></td>
                </tr>
                <%
					Dim i
					'sql ="SELECT DISTINCT tbl_PSystem_PurchaseOrder.PurOrderNo FROM  tbl_PSystem_PurchaseOrder INNER JOIN " & _
    	            '   " tbl_PSystem_Quotations ON tbl_PSystem_PurchaseOrder.PurOrderNo = tbl_PSystem_Quotations.PurOrderNo " & _
					'   " WHERE  (tbl_PSystem_Quotations.isClosed = 0)"

					sql = " SELECT DISTINCT a.PurOrderNo,b.isGRNEntered FROM  tbl_PSystem_PurchaseOrder a, tbl_PSystem_Quotations b " & _
						" where a.PurOrderNo = b.PurOrderNo and b.isClosed = 0 and b.isPOCancelled = 0 order by a.PurOrderNo "

					call RunSql(sql,rsItems)
					i = 1
					if rsItems.Eof = false then
					While Not  rsItems.EOF
						PurOrdNo =	rsItems("PurOrderNo")
						
							sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrdNo &" "
							call RunSql(sql,rsPONum)
							if rsPONum.Eof = false then
								PurOrderNum = rsPONum("PurOrderNum")
							end if
							rsPONum.Close

						if rsItems("isGRNEntered") = 0 then
							sStatus = "New PO Released"
						else
							sStatus = "Partial PO Received"
						end if
				%>
                <tr bgcolor="<%=gsBGColorLight%>">
                  <td align="center"><%=i%></td>
                  <td align="center"><a href="javascript:redirect(<%=PurOrdNo%>)"><%=GetPurchaseOrderNo(PurOrderNum)%></a></td>
                  <td align="center"><%=sStatus%></td>
                </tr>
                <%
					i = i + 1
					rsItems.movenext
					Wend
					else
						Response.Write "<tr ><td align='center' colspan='3'><br><b><font color='red'> No records found </font></b></td></tr>"
					end if
					rsItems.Close
				%>
                <tr class="blue">
                  <td align="center" colspan="3">&nbsp;</td>
                </tr>
              </Table>
            </form>
          </td>
        </tr>
      </table>
<p align="center">
<script language="javascript">
 function redirect(requestId)
 {
 	document.FinalForm.hdPurOrdNo.value = requestId;
	document.FinalForm.method="Post";
	document.FinalForm.action="PurOrderReleased_view.asp";
	document.FinalForm.submit();

 }

</script>
<form name="FinalForm">
	<input type="hidden" name="hdPurOrdNo" value="">
</form>
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->