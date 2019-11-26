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

<Script language=JavaScript src="../includes/javascript/validate.js" type=text/javascript></SCRIPT>
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
end if%>
<%
	ReceivedEditAction=Request.Form("EditAction")
	if ReceivedEditAction="" then
		ReceivedEditAction="Add"
	end if
	if ReceivedEditAction="Edit" then
		isReadOnly="readonly"
	else
		isReadOnly=""
	end if
	if ReceivedEditAction="Edit" then
			SupplierName=Request.Form("SupplierName")
			sql="sp_PSystem_EditSupplier '" & SupplierName & "'"
			call RunSQL(sql,rsSupplier)
			if not rsSupplier.eof then
				SupplierAddress=rsSupplier("SupplierAddress")
				ContactPerson=rsSupplier("ContactPerson")
				ContactPersonJobTitle=rsSupplier("ContactPersonJobTitle")
				TelephoneNo=rsSupplier("TelephoneNo")
				MobileNo=rsSupplier("MobileNo")
				EmailId=rsSupplier("EmailId")
				CSTNo=rsSupplier("CSTNo")
				TINNo=rsSupplier("TINNo")
				URL=rsSupplier("URL")
				TANNo=rsSupplier("TANNo")
				ServiceTaxNo=rsSupplier("ServiceTaxNo")
				isActive=rsSupplier("isActive")
				if isActive="True" then
					Chk = "Checked"
				else
					Chk = ""
				end if
			end if
			rsSupplier.close()
	end if
	if SupplierName = "" then
		SupplierName = Request("hdSupplier")
		'Response.write SupplierName & "supplierName"
	end if	
	%>
<table width="100%" cellspacing="2" cellpadding="2" border="0">
		<tr height="25" class="blue" align="center">
			<td width="40%" align="center">
				<font color=#ffffff><b>Admin--Suppliers</b></font>
			</td>
			<td width="60%" align="center">
				<font color=#ffffff><b>Workspace</b></font>
			</td>
		</tr>
		<tr >
			<td width="40%" align="center" valign="top">
				<table size="95%">
					<tr>
						<td>
							<!--#include file="SupplierList.asp"-->
						</td>
					</tr>
				</table>
			</td>
			<td width="60%" align="center" valign="top">
				<form name="SupplierForm">
					<table size="95%">
						<tr align="center" valign="middle">

                  <td colspan="2"> <b><%=ReceivedEditAction%> Supplier</b> &nbsp;(<font color="red"><b>*</b></font><b> 
                    Fields are mandatory</b> )</td>
						</tr>
						<tr align="center" valign="middle">

                  <td colspan="2">
				  <%
				  	if Request.Form("hdSup") = "" Then
						Response.write ""	
				  	elseif Request.Form("hdSup") = "Y" then
						Response.write "<font color='red'> Supplier already exist ! </font>" 
					elseif Request.Form("hdSup") = "N" then
						Response.write "<font color='green'>Supplier does not exist ! </font>"
					end if	
				  %>
				  </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Supplier
                    Name&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="SupplierName" size="40" maxlength="50" value="<%=SupplierName%>" <%=isReadOnly%>  onfocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; <a href="javascript:CheckSupplier();">Check Availability</a></td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Supplier
                    Address&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<textarea class="formstylemedium" name="SupplierAddress" rows="4" cols="31" onfocus="javascript:isToPropagate=true;"
								onKeyDown="textCounter(document.SupplierForm.SupplierAddress,document.SupplierForm.remLen3,255)"	onKeyUp="textCounter(document.SupplierForm.SupplierAddress,document.SupplierForm.remLen3,255)"><%=SupplierAddress%></textarea>
                    &nbsp;<b></b>&nbsp;<font color="red">Max (255 Chars)</font>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="255">
							</td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Contact
                    Person&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="ContactPerson" size="32" Maxlength="50" value="<%=ContactPerson%>"  onfocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Contact
                    Person's Job Title&nbsp;: &nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="ContactPersonJobTitle" size="32" Maxlength="50"  value="<%=ContactPersonJobTitle%>"onfocus="javascript:isToPropagate=true;">
							</td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Telephone
                    No.&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="TelephoneNo" size="32" Maxlength="20" value="<%=TelephoneNo%>"  onfocus="javascript:isToPropagate=true;">
                    &nbsp;<b></b>&nbsp; </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Mobile
                    No.&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="MobileNo" size="32" Maxlength="20"  value="<%=MobileNo%>" onfocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Email
                    Id&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="EmailId" size="32" Maxlength="50"  value="<%=EmailId%>" onfocus="javascript:isToPropagate=true;">
                    &nbsp;<b></b>&nbsp; </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>CST No.&nbsp;:
                    &nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="CSTNo" size="32" maxlength="15"  value="<%=CSTNo%>" onfocus="javascript:isToPropagate=true;">
                    <b></b> </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>TIN No.&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="TINNo" size="32" maxlength="15"  value="<%=TINNo%>" onfocus="javascript:isToPropagate=true;">
                    <b></b>&nbsp; </td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>URL&nbsp;:
                    &nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="URL" size="32" maxlength="50" value="<%=URL%>" onfocus="javascript:isToPropagate=true;">
							</td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>TAN No.&nbsp;:
                    &nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="TANNo" size="32" maxlength="10"  value="<%=TANNo%>" onfocus="javascript:isToPropagate=true;">
							</td>
						</tr>
						<tr height=25>

                  <td class="blue" align="Right"><font color=#ffffff><b>Service
                    Tax No.&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="text" name="ServiceTaxNo" size="32" maxlength="15"  value="<%=ServiceTaxNo%>" onfocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; </td>
						</tr>
						<% if ReceivedEditAction="Edit" then %>
						<tr height=25>
							<td class="blue" align="Right"><font color=#ffffff><b>Active Status&nbsp;:&nbsp;</b></font></td>
							<td bgcolor="<%=lclstr_bgColor%>">
								<input class="formstylemedium" type="checkbox" name="isActive" <%=Chk%>>
							</td>
						</tr>
						<% end if %>
						<tr align="center" valign="middle">
							<td colspan="2" bgcolor="<%=lclstr_bgColor%>" align="center">
								<input class="formbutton" type="button" value="Submit" style="border: 1 solid" onclick="SaveSupplier()">&nbsp;
								&nbsp;<input class="formbutton" type="Reset" value="Reset" style="border: 1 solid" >&nbsp;
							</td>
						</tr>
					</table>
					<input type="hidden" name="EditAction" value="<%=ReceivedEditAction%>">
				</form>
			</td>
		</tr>
	</table>
</form>
 <script language="javascript">
 function CheckSupplier()
 {
	var SupName = document.SupplierForm.SupplierName.value;
	document.FinalForm.hdSupName.value = SupName;

	document.FinalForm.method="Post";
	document.FinalForm.action="CheckSupplier.asp";
	document.FinalForm.submit();
	
 }
 </script>
 <form name="FinalForm">
	<input type="hidden" name="hdSupName" value="">
 </form>

<br>
<p align="center">
<a href="../../iMorfusAdmin/"><%=dictLanguage("Return_Admin_Home")%></a>
</p>
<script language="javascript">
	isToPropagate=true;
	function validateSupplierName(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value.toUpperCase());
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			ctrl.value=ctrl.value.replace(/PRIVATE/i,"PVT.");
			ctrl.value=ctrl.value.replace(/LIMITED/i,"LTD.");
			if(ctrl.value=="")
			{
				alert("Supplier Name is mandatory");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
	}
	function validateSupplierAddress(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(ctrl.value=="")
			{
				alert("Supplier Address is mandatory");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
	}
	function validateContactPerson(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(ctrl.value=="")
			{
				alert("Contact Person is mandatory");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
	}
	function validateContactPersonJobTitle(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			var regexp = new RegExp (/[a-zA-z]/);
			if(!regexp.test(ctrl.value))
			{
				alert("Enter a valid Contact Person Job title");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
	}
	function validateTelephoneNo(ctrl)
	{
		var regexp = new RegExp (/([\(\+])?([0-9]{1,3}([\s])?)?([\+|\(|\-|\)|\s])?([0-9]{2,4})([\-|\)|\.|\s]([\s])?)?([0-9]{2,4})?([\.|\-|\s])?([0-9]{4,8})/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		if(isToPropagate==true)
		{
			if(ctrl.value=="")
			{
				alert("Telephone number is mandatory");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else if(!regexp.test(ctrl.value))
			{
				alert("Enter a valid telephone number");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;
		}
		return true;
	}
	function validateMobileNo(ctrl)
	{
		var regexp = new RegExp (/([\(\+])?([0-9]{1,3}([\s])?)?([\+|\(|\-|\)|\s])?([0-9]{2,4})([\-|\)|\.|\s]([\s])?)?([0-9]{2,4})?([\.|\-|\s])?([0-9]{4,8})/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		if(isToPropagate==true)
		{
			if(ctrl.value=="")
			{
				alert("Mobile number is mandatory");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else if(!regexp.test(ctrl.value))
			{
				alert("Enter a valid mobile number");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;
		}
		return true;
	}
	function validateEmailId(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space

			if(ctrl.value=="")
			{
				alert("Email Id is mandatory");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else if(!isEmail(ctrl.value))
			{
				alert("Please enter valid email id");
				ctrl.focus();
				return false;
			}
 /*			else if (ctrl.value.length >0) 
			{
				 i=ctrl.value.indexOf("@")
				 j=ctrl.value.indexOf(".",i)
				 k=ctrl.value.indexOf(",")
				 kk=ctrl.value.indexOf(" ")
				 jj=ctrl.value.lastIndexOf(".")+1
				 len=ctrl.value.length

			if ((i>0) && (j>(1+1)) && (k==-1) && (kk==-1) && (len-jj >=2) && (len-jj<=3)) 
			{
			}
		 	else 
			{
		 		alert("Please enter a valid email address");
				ctrl.focus();
				return false;
			}

 			}
	*/
			else
			{
				isToPropagate=true;
				return true;
			}
		}
	}
	function validateTINNo(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space

			var regexp = new RegExp (/^[0-9]\d*$/);
			if(!regexp.test(ctrl.value))
			{
				alert("Enter a valid TIN Number");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
    }
	function validateURL(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space

			if(ctrl.value!="")
			{
			var regexp = /(\w+:{0,1}\w*@)?(\S+)(:[0-9]+)?(\/|\/([\w#!:.?+=&%@!\-\/]))?/
			if(!regexp.test(ctrl.value))
			{
				alert("Enter a valid URL");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
			}
			else
			{
				isToPropagate=true;
				return true;
			}

		}
    }
	function validateCSTNo(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space

			if(ctrl.value!="")
			{
				var regexp = new RegExp (/^[0-9]\d*$/);
				if(!regexp.test(ctrl.value))
				{
					alert("Enter a valid CST Number");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
				{
					isToPropagate=true;
					return true;
				}
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
    }

	function validateTANNo(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space

			if(ctrl.value!="")
			{
				var regexp = new RegExp (/^[0-9]\d*$/);
				if(!regexp.test(ctrl.value))
				{
					alert("Enter a valid TAN Number");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
				{
					isToPropagate=true;
					return true;
				}
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
    }
   	function validateServiceTaxNo(ctrl)
	{
		if(isToPropagate==true)
		{
			ctrl.value=trim(ctrl.value);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space

			var regexp = new RegExp (/^[0-9]\d*$/);
			if(!regexp.test(ctrl.value))
			{
				alert("Enter a valid Service Tax number");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			{
				isToPropagate=true;
				return true;
			}
		}
    }

	function SaveSupplier()
	{
		if(validateSupplierName(document.SupplierForm.SupplierName)==false)
			return false;
		if(validateSupplierAddress(document.SupplierForm.SupplierAddress)==false)
			return false;
		if(validateContactPerson(document.SupplierForm.ContactPerson)==false)
			return false;
		if(validateTelephoneNo(document.SupplierForm.TelephoneNo)==false)
			return false;
		if(validateMobileNo(document.SupplierForm.MobileNo)==false)
			return false;
		if(validateEmailId(document.SupplierForm.EmailId)==false)
			return false;
		if(validateCSTNo(document.SupplierForm.CSTNo)==false)
			return false;
		if(validateTINNo(document.SupplierForm.TINNo)==false)
			return false;
		if(validateURL(document.SupplierForm.URL)==false)
			return false;
		if(validateTANNo(document.SupplierForm.TANNo)==false)
			return false;
		if(validateServiceTaxNo(document.SupplierForm.ServiceTaxNo)==false)
			return false;

		document.SupplierForm.method="post";
		document.SupplierForm.action="SupplierEdit.asp"
		document.SupplierForm.submit();
	}
</script>
<SCRIPT LANGUAGE="JavaScript">

<!-- Web Site:  The JavaScript Source -->
<!-- Use one function for multiple text areas on a page -->
<!-- Limit the number of characters per textarea -->
<!-- Begin
function textCounter(field,cntfield,maxlimit) {
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);

// otherwise, update 'characters left' counter
else
cntfield.value = maxlimit - field.value.length;
}
//  End -->
</script>
<!--#include file="../includes/main_page_close.asp"-->