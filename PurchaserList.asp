<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Subhash Rampuri
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>


<form name="strFormm">
	<table cellspacing="2" cellpadding="2" width="275" border="0">
		<tr  height="25" class="blue">
			<td width="70%" align="center"><font color=#ffffff><b>Purchaser Name</b></font></td>
			<td align="center"><font color=#ffffff><b>Employee Id</b></font></td>
		</tr>
		<%
			lclstr_bgColor = gsBGColorLight
			sql = sql_GetPurchasersByPriority
			call runSQL(sql,rsPurchasers)
			while not rsPurchasers.eof
			if lclstr_bgColor = gsBGColorLight then
				lclstr_bgColor = gsBGColorDark
			else
				lclstr_bgColor = gsBGColorLight
			end if
		%>
		<tr height="25">
			<td bgcolor="<%=lclstr_bgColor%>"><input type="checkbox" name="Purchaser" value="<%=rsPurchasers("EmployeeId")%>"><%=rsPurchasers("EmployeeName")%></input></td>
			<td bgcolor="<%=lclstr_bgColor%>"><%=rsPurchasers("EmployeeId")%></td>
		</tr>
		<%
			rsPurchasers.movenext
			wend
			rsPurchasers.close

			if lclstr_bgColor = gsBGColorLight then
				lclstr_bgColor = gsBGColorDark
			else
				lclstr_bgColor = gsBGColorLight
			end if

		%>
		<tr>
			<td colspan=2 align="center" bgcolor="<%=lclstr_bgColor%>">
				<input class="formbutton" type="button" name="Up" value="   Top  " style="border: 1 solid" onclick="javascript:MoveUp();"></input>&nbsp;
				<input class="formbutton" type="button" name="Down" value=" Down " style="border: 1 solid" onclick="javascript:MoveDown();"></input>&nbsp;
				<input class="formbutton" type="button" name="Delete" value="Delete" style="border: 1 solid" onclick="javascript:DeletePurchaser();"></input>
			</td>
		</tr>
	</table>
	<input type="hidden" name="EditAction" value="">
	<input type="hidden" name="Purchaser" value="">
</form>

<script language="javascript">
	function MoveUp()
	{
	  if(CheckSelection())
	  {
		  document.strFormm.EditAction.value="MoveUp";
		  document.strFormm.method="post";
		  document.strFormm.action="PurchaserEdit.asp";
		  document.strFormm.submit();
	  }
	  else
	  	alert("Please select Purchaser");
	}
	function MoveDown()
	{
	  if(CheckSelection())
	  {
		  document.strFormm.EditAction.value="MoveDown";
		  document.strFormm.method="post";
		  document.strFormm.action="PurchaserEdit.asp";
		  document.strFormm.submit();
	  }
	  else
	  	alert("Please select Purchaser");
	}
	function DeletePurchaser()
	{
	  if(CheckSelection())
	  {
		  document.strFormm.EditAction.value="Delete";
		  document.strFormm.method="post";
		  document.strFormm.action="PurchaserEdit.asp";
		  document.strFormm.submit();
	  }
	  else
	  	alert("Please select Purchaser");
	}
	function CheckSelection()
	{
		var count=document.strFormm.Purchaser.length;
		var selcount=0;
		for(i=0;i<count;i++)
			if(document.strFormm.Purchaser[i].checked)
				selcount++;
		if(selcount>0)
			return true;
		else
			return false;
	}
</script>

