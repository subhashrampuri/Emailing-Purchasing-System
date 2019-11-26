<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Vinay Kumar
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>
<!--#include file="../includes/SiteConfig.asp"-->
<!--#include file="../includes/Connection_open.asp"-->
<%
	select case Request.Form("EditAction")
		case "Add"
		sql="sp_PSystem_AddSupplier '" & Replace(Server.HTMLEncode(Request.Form("SupplierName")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("SupplierAddress")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("ContactPerson")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("ContactPersonJobTitle")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("TelephoneNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("MobileNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("EmailId")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("CSTNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("TINNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("URL")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("TANNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("ServiceTaxNo")),"'","''") & "'"
		Call DoSQL(sql)
		'Response.write sql	
		case "Edit"
		'	sql="select SupplierName from tbl_PSystem_Supplier"
		'	call RunSql(sql,rsSupplier)
		'	Response.write rsSupplier.eof
		'	if not rsSupplier.eof then
		'		while not rsSupplier.eof
		'			Response.write rsSupplier("SupplierName") & "<br>"
		'			rsSupplier.movenext
		'		wend
		'	end if
			if Request.Form("isActive")="on" then
				isChecked=1
			else
				isChecked=0
			end if
	sql="sp_PSystem_SaveEditSupplier '" & Replace(Server.HTMLEncode(Request.Form("SupplierName")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("SupplierAddress")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("ContactPerson")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("ContactPersonJobTitle")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("TelephoneNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("MobileNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("EmailId")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("CSTNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("TINNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("URL")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("TANNo")),"'","''") & "','" & Replace(Server.HTMLEncode(Request.Form("ServiceTaxNo")),"'","''") & "'," & isChecked
	'Response.write sql
	call DoSQL(sql)
	end select
%>
<form name="editForm">
	<input type="hidden" name="EditAction" value="">
	<input type="hidden" name="SupplierName" value="">
	<input type="hidden" name="SupplierExist" value="">
</form>
<script language="javascript">
	<%
		SupplierName=Request.Form("SupplierName")
	%>
	document.editForm.EditAction.value="<%=EditAction%>";
	document.editForm.SupplierName.value="<%=SupplierName%>";
	//document.editForm.SupplierExist.value="<%=Flag%>";
	document.editForm.method="post";
	document.editForm.action="Suppliers.asp"
	document.editForm.submit();
</script>