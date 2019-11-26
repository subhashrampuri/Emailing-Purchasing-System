<%@LANGUAGE="VBSCRIPT"%>
<%Response.Expires=-1%>
<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Vinay Kumar
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>

<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<%
	EmployeeId=Session("Employee_Id")

	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	'Check whether the logged in person is a assigned approver in Purchase System
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	sql="select EmployeeId from tbl_PSystem_Approver where EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsPSystemApprover)
	if not rsPSystemApprover.eof then
		isPSystemApprover=true
	else
		isPSystemApprover=false
	end if
	rsPSystemApprover.close()
	''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
	
	if not (isPSystemApprover=true) then
		Response.write "<center><br><br><br><br><br><br><br><br><br><br><br>"
		Response.write "<font color=Red>You are not authorized to view this page.</font></center>"
		Response.end
	end if
%>

	<table width="100%" align="center" >
	<tr class="blue">
	      <td align="center"><font color="#ffffff"><b>Approvers - Inbox</b></font></td>
	</tr>

	<tr>
	<td align="center">&nbsp;</td>
	</tr>
	<tr>
	<td align="center">&nbsp;</td>
	</tr>
	<tr>
	<td align="center">
            <table border="0" cellspacing="2" cellpadding="2">
              <tr bgcolor="<%=gsBGColorLight%>">
                <td align="center"><a href ="PurchaseRequestList.asp" style="text-decoration: none"><b>Purchase
                  Requests For Approval</b></a></td>
              </tr>
              <tr bgcolor="<%=gsBGColorLight%>">
                <td align="center"><a href ="Purchase_QuotationList.asp" style="text-decoration: none"><b>Purchase
                  Quotations For Approval</b></a></td>
              </tr>
            </table>
          </td>
	</tr>
	</table>
<br>
<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->