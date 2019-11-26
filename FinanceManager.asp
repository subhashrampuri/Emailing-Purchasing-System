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

<script LANGUAGE="JavaScript" SRC="../../includes/javascript/validate.js"></script>
	<table width="100%" cellspacing="2" cellpadding="2" border="0">
		<tr height="25" class="blue">
			<td width="40%" align="center">
				<font color=#ffffff><b>Admin--Finance Managers</b></font>
			</td>
			<td width="60%" align="center">
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
						<td>
							<!--#include file="FinanceManagerList.asp"-->
						</td>
					</tr>
				</table>
			</td>
			<td width="60%" valign="top">
				<form name="FinanceManagerForm">
					<table size="95%" align="center">
						<tr align="center">
							
                  <td colspan="2"> <b>Add Finance Managers</b></td>
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
							sql=sql_GetCandidatesForFinanceManager()
							call RunSql(sql,rsFinanceManagerCandidates)
							lclstr_bgColor=gsBGColorLight
						%>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Employee&nbsp;:&nbsp;<font color="red">*</font></b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<select class="formstyleMedium" name="FinanceManager" style="border: 1 solid">
									<option value=0 selected>&nbsp;&nbsp;--Select&nbsp;Finance Managers--</option>
									<%
										While not rsFinanceManagerCandidates.eof
									%>
									<option value="<%=rsFinanceManagerCandidates("EmployeeId")%>"><%=rsFinanceManagerCandidates("EmployeeName")%> : <%=rsFinanceManagerCandidates("EmployeeId")%></option>
									<%
										rsFinanceManagerCandidates.MoveNext()
										wend
										rsFinanceManagerCandidates.close()
										set rsFinanceManagerCandidates=nothing
									%>
								</select>
							</td>
						</tr>
						<tr align="center" valign="middle">
							<td colspan="2" bgcolor="<%=lclstr_bgColor%>">
								<input class="formbutton" type="button" name="ApproverAdd" value="Submit" style="border: 1 solid" onclick="javascript:UpdateFinanceManager();">&nbsp;
							</td>
						</tr>
					</table>
					<input type="hidden" name="EditAction" value="">
				</form>
			</td>
		</tr>
	</table>

<br>
<p align="center">
<a href="../../iMorfusAdmin/"><%=dictLanguage("Return_Admin_Home")%></a>
</p>
<script language="javascript">
	function UpdateFinanceManager()
	{
		if(document.FinanceManagerForm.FinanceManager.value==0)
			alert("Employee is a Required field");
		else
		{
			document.FinanceManagerForm.EditAction.value="Add";
			document.FinanceManagerForm.method="post";
			document.FinanceManagerForm.action="FinanceManagerEdit.asp";
			document.FinanceManagerForm.submit();
		}
	}
</script>
<!--#include file="../includes/main_page_close.asp"-->




