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
	
  <table cellspacing="0" cellpadding="0" width="275" border="0">
    <tr> 
      <td align="right" colspan="2"> 
        <input type="button" class="formbutton" value=" Add  " style="border: 1 solid" onclick="javascript:AddSupplier()">
      </td>
    </tr>
    <tr  height="25" class="blue"> 
      <td><p style="margin-left:10"><font color=#ffffff><b>Supplier Name</b></font></p></td>
      <td>
        <p style="margin-right:10"><font color=#ffffff><b>Action</b></font>
      </td>
    </tr>
    <%
			lclstr_bgColor = gsBGColorLight
			sql = sql_GetSuppliersByName()
			call runSQL(sql,rsSuppliers)
			if not rsSuppliers.eof then
			while not rsSuppliers.eof
				if lclstr_bgColor = gsBGColorLight then
					lclstr_bgColor = gsBGColorDark
				else
					lclstr_bgColor = gsBGColorLight
				end if
		%>
    <tr height="25"> 
      <td bgcolor=<%=lclstr_bgColor%> ><p style="margin-left:10"><%=rsSuppliers("SupplierName")%></p></td>
      <td bgcolor=<%=lclstr_bgColor%> width=15%><p style="margin-right:10"><a value="<%=rsSuppliers("SupplierName")%>" style="CURSOR: hand" onclick="javascript:EditSupplier(this.value);">Edit</a></p></td>
    </tr>
    <%
				rsSuppliers.movenext
			wend
			end if
			rsSuppliers.close

			if lclstr_bgColor = gsBGColorLight then
				lclstr_bgColor = gsBGColorDark
			else
				lclstr_bgColor = gsBGColorLight
			end if

		%>
  </table>
	<input type="hidden" name="EditAction" value="">
	<input type="hidden" name="SupplierName" value="">
</form>

<script language="javascript">
	function EditSupplier(val)
	{
		document.strFormm.EditAction.value="Edit"
		document.strFormm.SupplierName.value=val;
		document.strFormm.method="post";
		document.strFormm.action="Suppliers.asp"
		document.strFormm.submit();
	}
	function AddSupplier()
	{
		document.strFormm.EditAction.value="Add"
		document.strFormm.method="post";
		document.strFormm.action="Suppliers.asp"
		document.strFormm.submit();
	}
</script>

