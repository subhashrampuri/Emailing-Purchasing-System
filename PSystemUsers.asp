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
EmployeeId=Session("Employee_Id")

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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'****************************************************************************************************************************************************
if not (isSiteAdmin=true or (isPSystemUser=true and PSystemAccessLevel="Administrator")) then
	Response.write "<center><br><br><br><br><br><br><br><br><br><br><br>"
	Response.write "<font color=Red>You are not authorized to view this page.</font></center>"
	Response.end
end if
%>
<script LANGUAGE="JavaScript" SRC="../../includes/javascript/validate.js"></script>
	<table width="100%" cellspacing="2" cellpadding="2" border="0">
		<tr height="25" class="blue">
			<td width="50%" align="center">
				<font color=#ffffff><b>Purchase System Extended Users</b></font>
			</td>
			<td width="50%" align="center">
				<font color=#ffffff><b>Workspace</b></font>
			</td>
		</tr>
		<tr >
			<td width="50%" valign="top">
				<table width="100%" align="center" valign="top">
					<tr>
						<td>
							&nbsp;
						</td>
					</tr>
					<tr>
						<td>
							<!--#include file="PSystemUsersList.asp"-->
						</td>
					</tr>
				</table>
			</td>
			<td width="50%" valign="top">
				<form name="UserForm">
					<table size="95%" align="center">
						<tr align="center">
							<td colspan="2">
								<b>Add User</b>
							</td>
						</tr>
						<tr align="center">
							<td colspan="2">
								<b><font color="red">*</font>&nbsp;Fields are mandatory</b>
							</td>
						</tr>
						<tr>
							<td colspan="2">
								&nbsp;
							</td>
						</tr>
						<%
							sql=sql_GetCandidatesForPSystemAccess()
							call RunSql(sql,rsPSystemAccessCandidates)
							lclstr_bgColor=gsBGColorLight
						%>
						<tr height=25>
							<td class="blue" align="Right"><font color=#ffffff><b>Select Employee&nbsp;:&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<select class="formstyleMedium" name="User" style="border: 1 solid">
									<option value=0 selected>-----Employee-----</option>
									<%
										While not rsPSystemAccessCandidates.eof
									%>
									<option value="<%=rsPSystemAccessCandidates("EmployeeId")%>"><%=rsPSystemAccessCandidates("EmployeeName")%> : <%=rsPSystemAccessCandidates("EmployeeId")%></option>
									<%
										rsPSystemAccessCandidates.MoveNext()
										wend
										rsPSystemAccessCandidates.close()
										set rsPSystemAccessCandidates=nothing
									%>
								</select>
							</td>
						</tr>
						<%
							sql=sql_GetPSystemAccessLevels()
							call RunSQL(sql,rsAccessLevels)
						%>
						<tr height=25>
							<td class="blue" align="Right"><font color=#ffffff><b>Select Access Level&nbsp;:&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<select class="formstyleMedium" name="AccessLevel" style="border: 1 solid">
									<option value=0 selected>-----Access Level-----</option>
									<%while not rsAccessLevels.eof
										AccessLevel=rsAccessLevels("AccessLevel")
										AccessLevelText=rsAccessLevels("AccessLevelText")
									%>
									<option value=<%=AccessLevel%>><%=AccessLevelText%></option>
									<%
										rsAccessLevels.movenext
									wend
									%>
								</select>
							</td>
						</tr>
						<tr align="center" valign="middle">
							<td colspan="2" bgcolor="<%=lclstr_bgColor%>">
								<input class="formbutton" type="button" name="UserAdd" value="Submit" style="border: 1 solid" onclick="javascript:AddAdmin();"></input>&nbsp;
							</td>
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
<a href="../../imorfusadmin/"><%=dictLanguage("Return_Admin_Home")%></a>
</p>
<script language="javascript">
	function AddAdmin()
	{
		if(document.UserForm.User.value==0)
			alert("Please select employee");
		else if(document.UserForm.AccessLevel.value==0)
			alert("Please select access level");
		else
		{
			document.UserForm.EditAction.value="Add";
			document.UserForm.method="post";
			document.UserForm.action="PSystemUsersEdit.asp";
			document.UserForm.submit();
		}
	}
</script>
<!--#include file="../includes/main_page_close.asp"-->




