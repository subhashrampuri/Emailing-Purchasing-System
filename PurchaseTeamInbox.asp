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
	<table width="100%" align="center" >
	<%
		EmployeeId=Session("Employee_Id")
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		'Check whether the logged in person is a assigned purchaser in Purchase System
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		sql="select EmployeeId from tbl_PSystem_PurchaseTeam where EmployeeId='" & EmployeeId & "'"
		call RunSQL(sql,rsPSystemPurchaser)
		if not rsPSystemPurchaser.eof then
			isPSystemPurchaser=true
		else
			isPSystemPurchaser=false
		end if
		rsPSystemPurchaser.close()
		''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
		if not (isPSystemPurchaser=true) then
			Response.write "<center><br><br><br><br><br><br><br><br><br><br><br>"
			Response.write "<font color=Red>You are not authorized to view this page.</font></center>"
			Response.end
		end if

	%>
	<tr class="blue">
	      <td align="center"><font color="#ffffff"><b>Purchase Team - Inbox</b></font></td>
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
                <td align="center"><a href ="PurchaseApproveList.asp" style="text-decoration: none"><b>Approved Purchase Requests</b></a></td>
              </tr>
              <tr bgcolor="<%=gsBGColorLight%>">
                <td align="center"><a href ="PurchaseOrder.asp" style="text-decoration: none"><b>Approved Purchase Quotations</b></a></td>
              </tr>
              <tr bgcolor="<%=gsBGColorLight%>">
                <td align="center"><a href="PurchaseOrderReleased.asp" style="text-decoration:none"><b>Released Purchase Orders</b></a></td>
              </tr>
              <tr bgcolor="<%=gsBGColorLight%>">
                <td align="center"><a href="Purchase_GRN.asp" style="text-decoration:none"><b>Goods
                  Received Note</b></a></td>
              </tr>
              <tr bgcolor="<%=gsBGColorLight%>">
                <td align="center"><a href="GRNCloseList.asp" style="text-decoration:none"><b>Close Goods Received Note</b></a></td>
              </tr>
            </table>
          </td>
	</tr>
	</table>
<br>
<br>
<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->