<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Vinay Kumar
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>


<form name="strFormm">
	<table cellspacing="2" cellpadding="2" width="95%" border="0" align="center">
		<tr  height="25" class="blue">
			<td width="60%" align="center"><font color=#ffffff><b>User Name</b></font></td>
			<td width="15%" align="center"><font color=#ffffff><b>Employee Id</b></font></td>
			<td width="25%" align="center"><font color=#ffffff><b>Access Level</b></font></td>
		</tr>
		<%
			lclstr_bgColor = gsBGColorLight
			sql = sql_GetPSystemUsers
			call runSQL(sql,rsUsers)
			while not rsUsers.eof
			if lclstr_bgColor = gsBGColorLight then
				lclstr_bgColor = gsBGColorDark
			else
				lclstr_bgColor = gsBGColorLight
			end if
		%>
		<tr height="25">
			<td bgcolor="<%=lclstr_bgColor%>"><input type="checkbox" name="User" value="<%=rsUsers("EmployeeId")%>"><%=rsUsers("EmployeeName")%></input></td>
			<td bgcolor="<%=lclstr_bgColor%>" align="center"><%=rsUsers("EmployeeId")%></td>
			<td bgcolor="<%=lclstr_bgColor%>" align="center"><%=rsUsers("AccessLevelText")%></td>
		</tr>
		<%
			rsUsers.movenext
			wend
			rsUsers.close

			if lclstr_bgColor = gsBGColorLight then
				lclstr_bgColor = gsBGColorDark
			else
				lclstr_bgColor = gsBGColorLight
			end if
		%>
		<tr>
			<td colspan=3 align="right" bgcolor="<%=lclstr_bgColor%>">
				<input class="formbutton" type="button" name="Delete" value="Delete" style="border: 1 solid" onclick="javascript:DeleteUser();"></input>
			</td>
		</tr>
	</table>
	<input type="hidden" name="EditAction" value="">
	<input type="hidden" name="User" value="">
</form>

<script language="javascript">
	function DeleteUser()
	{
	  if(CheckSelection())
	  {
		  document.strFormm.EditAction.value="Delete";
		  document.strFormm.method="post";
		  document.strFormm.action="PSystemUsersEdit.asp";
		  document.strFormm.submit();
	  }
	  else
	  	alert("Please select User");
	}

	function CheckSelection()
	{
		var count=document.strFormm.User.length;
		var selcount=0;
		for(i=0;i<count;i++)
			if(document.strFormm.User[i].checked)
				selcount++;
		if(selcount>0)
			return true;
		else
			return false;
	}

</script>

