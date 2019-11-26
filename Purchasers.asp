<%@LANGUAGE="VBSCRIPT"%>
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
EmployeeId=Session("employee_Id")
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Check whether the logged in person is Finance Manager
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql="select FinanceManager from tbl_PSystem_FinanceManager where FinanceManager='" & EmployeeId & "'"
call RunSQL(sql,rsFinanceManager)
if not rsFinanceManager.eof then
	isFinanceManager=true
else
	isFinanceManager=false
end if
rsFinanceManager.close()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Check whether the logged in person is a site admin
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql="select EmployeeId from tbl_SiteAdmin where EmployeeId='" & EmployeeId & "'"
call RunSQL(sql,rsSiteAdmin)
if not rsSiteAdmin.eof then
	isSiteAdmin=true
else
	isSiteAdmin=false
end if
rsSiteAdmin.close()
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Check whether the logged in person is an extended user of Purchase System
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql="select EmployeeId,AccessLevelText from tbl_PSystem_Users A,tbl_PSystem_UserAccessLevel B where A.AccessLevel=B.AccessLevel and EmployeeId='" & EmployeeId & "'"
call RunSQL(sql,rsPSystemUser)
if not rsPSystemUser.eof then
	isPSystemUser=true
	PSystemAccessLevel=rsPSystemUser("AccessLevelText")
else
	isPSystemUser=false
	PSystemAccessLevel=""
end if
rsPSystemUser.close()
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''if isSiteAdmin=true or (isPSystemUser=true and PSystemAccessLevel="Administrator")=true or isFinanceManager=true then
if not (isSiteAdmin=true or (isPSystemUser=true and PSystemAccessLevel="Administrator")=true or isFinanceManager=true) then
	Response.write "<center><br><br><br><br><br><br><br><br><br><br><br>"
	Response.write "<font color=Red>You are not authorized to view this page.</font></center>"
	Response.end
end if
%>

<script LANGUAGE="JavaScript" SRC="../../includes/javascript/validate.js"></script>
	<table width="100%" cellspacing="2" cellpadding="2" border="0">
		<tr height="25" class="blue" align="center">
			<td width="40%">
				<font color=#ffffff><b>Admin--Purchase Team</b></font>
			</td>
			<td width="60%">
				<font color=#ffffff><b>Workspace</b></font>
			</td>
		</tr>
		<tr >
			<td width="40%" valign="top">
				<table size="95%" align="center" valign="top">
					<tr>
						<td>&nbsp;

						</td>
					</tr>
					<tr>
						<td align="center">
							<!--#include file="PurchaserList.asp"-->
						</td>
					</tr>
				</table>
			</td>
			<td width="60%" valign="top">
				<form name="PurchaserForm">
					<table size="95%" align="center">
						<tr align="center">
							<td colspan="2">
								<b>Add Purchaser</b>
							</td>
						</tr>
						<tr align="center">
							<td colspan="2">
								<b><font color="red">*</font>&nbsp;Fields are mandatory</b>
							</td>
						</tr>
						<tr>
							<td colspan="2">&nbsp;

							</td>
						</tr>
						<%
							sql=sql_GetCandidatesForPurchaser()
							call RunSql(sql,rsPurchaseCandidates)
							lclstr_bgColor=gsBGColorLight
						%>
						<tr height=25>
							<td class="blue" align="Right"><font color=#ffffff><b>Employee&nbsp;:&nbsp;<font color="red">*</font></b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<select class="formstyleMedium" name="Purchaser" style="border: 1 solid">
									<option value=0 selected>--Select&nbsp;Purchaser--</option>
									<%
										While not rsPurchaseCandidates.eof
									%>
									<option value="<%=rsPurchaseCandidates("EmployeeId")%>"><%=rsPurchaseCandidates("EmployeeName")%> : <%=rsPurchaseCandidates("EmployeeId")%></option>
									<%
										rsPurchaseCandidates.MoveNext()
										wend
										rsPurchaseCandidates.close()
										set rsPurchaseCandidates=nothing
									%>
								</select>
							</td>
						</tr>
						<tr align="center" valign="middle">
							<td colspan="2" bgcolor="<%=lclstr_bgColor%>">
								<input class="formbutton" type="button" name="PurchaserAdd" value="Submit" style="border: 1 solid" onclick="javascript:AddPurchaser();">
                    &nbsp; </td>
						</tr>
					</table>
					<input type="hidden" name="EditAction" value="">
				</form>
			</td>
		</tr>
	</table>
</form>
<br>
<p align="center">
<a href="../../iMorfusAdmin/"><%=dictLanguage("Return_Admin_Home")%></a>
</p>
<script language="javascript">
	function AddPurchaser()
	{
		if(document.PurchaserForm.Purchaser.value==0)
			alert("Employee is a Required field");
		else
		{
			document.PurchaserForm.EditAction.value="Add";
			document.PurchaserForm.method="post";
			document.PurchaserForm.action="PurchaserEdit.asp";
			document.PurchaserForm.submit();
		}
	}
</script>
<!--#include file="../includes/main_page_close.asp"-->




