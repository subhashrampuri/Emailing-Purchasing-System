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
ReceivedEditAction="Add"
EmployeeId=Session("Employee_Id")

'****************************************************************************************************************************************************
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Check whether the logged in person is a Manager or not.
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sqlvalid = "select EmployeeID from VSPL_Managers where EmployeeID = '" & EmployeeId & "'"
call RunSQL(sqlvalid,rsValid)
if not rsValid.eof then
	MngrEmpId = rsValid("EmployeeID")
end if
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'Get department of the logged in person
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
sql="select DeptName from VSPL_EmployeeHREntry A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
call RunSQL(sql,rsDepartment)
if not rsDepartment.eof then
	EmployeeDepartment=trim(rsDepartment("DeptName"))
else
	EmployeeDepartment=""
end if
rsDepartment.close()
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
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'****************************************************************************************************************************************************

if not (MngrEmpId=EmpId or EmployeeDepartment="ITIT" or EmployeeDepartment="HR" or EmployeeDepartment="Administration" or EmployeeDepartment="Operations" or (isPSystemUser=true and PSystemAccessLevel="General User")=true) then
	Response.write "<center><br><br><br><br><br><br><br><br><br><br><br>"
	Response.write "<font color=Red>You are not authorized to view this page.</font></center>"
	Response.end
end if

%>
<table width="100%" cellspacing="2" cellpadding="2" border="0">
		<tr height="25"  align="center">
			<td align="center">
				<font ><b>Manager--Purchase Request</b></font>
			</td>
		</tr>
		<tr >
			<td align="center" valign="top">
				<form name="strFormm">
					<table size="95%">
						<tr align="center" valign="middle">

                  <td colspan="6"> <b><%=ReceivedEditAction%>&nbsp;Item</b>&nbsp;&nbsp;</td>
						</tr>
						<tr align="center" valign="middle">

                  <td colspan="6"> <b>&nbsp;&nbsp;<font color="red">*</font>&nbsp;Fields
                    are mandatory</b></td>
						</tr>

						<%
							if MngrEmpId=EmployeeId or EmployeeDepartment="ITIT" or EmployeeDepartment="HR" or EmployeeDepartment="Administration" or EmployeeDepartment="Operations" then
								sql="select ProjectId,ProjectName from VSPL_PM_ProjectDetails A where (A.Status like 'On Going' or A.Status like 'On Hold' or A.Status like 'Expected') order by ProjectName"
							else
								sql="select A.ProjectId,A.ProjectName from VSPL_PM_ProjectDetails A, VSPL_Managers B where A.ManagerId=B.ManagerId and (A.Status like 'On Going' or A.Status like 'On Hold' or A.Status like 'Expected') and B.EmployeeId='" & EmployeeId & "' order by A.ProjectName"
							end if
							call RunSQL(sql,rsProjects)
						%>
						<tr height=25>

                  <td  align="Right"><font ><b>Project&nbsp;:<font color="red">*</font></b></font></td>
							<td  colspan="2">
								<select  name="Project"  onfocus="javascript:isToPropagate=true;">
									<option value="0" selected>--Select&nbsp;Project--</option>
								<%
								if MngrEmpId=EmployeeId or EmployeeDepartment="ITIT" or EmployeeDepartment="HR" or EmployeeDepartment="Administration" or EmployeeDepartment="Operations" then
										sql="select ProjectId,ProjectName from VSPL_PM_ProjectDetails A where (A.Status like 'On Going' or A.Status like 'On Hold' or A.Status like 'Expected') order by ProjectName"
									else
										sql="select A.ProjectId,A.ProjectName from VSPL_PM_ProjectDetails A, VSPL_Managers B where A.ManagerId=B.ManagerId and (A.Status like 'On Going' or A.Status like 'On Hold' or A.Status like 'Expected') and B.EmployeeId='" & EmployeeId & "' order by A.ProjectName"
									end if
									Call RunSQL(sql,rsProjects)
									if not rsProjects.eof then
										while not rsProjects.eof
											ProjectId=rsProjects("Projectid")
											ProjectName=rsProjects("ProjectName")
								%>
											<option value="<%=ProjectId%>"><%=ProjectName%></option>
								<%
									rsProjects.movenext
									wend
									end if
									rsProjects.close()
								%>
								</select>
                    &nbsp;&nbsp; </td>

                  <td  align="Right"><font ><b>Item Description&nbsp;:<font color="red">*</font></b></font></td>
							<td  colspan="2">
								<input  type="text" name="ItemDescription" size="40" maxlength="50"  onfocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; </td>
						</tr>
						<tr height=25>

                  <td  align="Right"><font ><b>Quantity
                    Required&nbsp;:<font color="red">*</font></b></font></td>
							<td  colspan="2">
								<input class="formstylemedium" type="text" name="QuantityRequired" size="11" maxlength="4"  onfocus="javascript:isToPropagate=true;">
                    &nbsp; </td>

                  <td  align="Right"><font ><b>Purchase/Service&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td >
								<input class="formstylemedium" type="radio" name="PurchaseOrService" value="0" text='Purchase' checked>Purchase
							</td>
							<td >
								<input class="formstylemedium" type="radio" name="PurchaseOrService" value="1" text='Service'>
                    Service&nbsp;&nbsp; </td>
						</tr>
						<tr height=25>

                  <td  align="Right"><font ><b>Required
                    Date&nbsp;: <font color="red">*</font></b></font></td>
							<td  colspan="2">
								<input class="formstylemedium" type="text" name="RequiredDate" size="11" value="" onfocus="javascript:isToPropagate=true;" Readonly>
                    &nbsp;<a name="CalanderLink" onClick="openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300)"><img border="0" src="/gif/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" width="16" height="15"></a>&nbsp;&nbsp;
                  </td>

                  <td  align="Right"><font ><b>Currency&nbsp;:
                    &nbsp;</b></font></td>
							<td >
								<input class="formstylemedium" type="radio" name="RupeeOrDollar" value="Rupee(s)" checked>
                    Rupees(Rs.)</td>
							<td >
								<input class="formstylemedium" type="radio" name="RupeeOrDollar" value="Dollar(s)">
                    Dollars($) </td>
						</tr>
						<tr height=25>

                  <td  align="Right"><font ><b>Purpose&nbsp;:<font color="red">*</font>&nbsp;</b></font></td>
							<td  colspan="2">
								<textarea class="formstylemedium" name="Purpose" rows="4" cols="30"  onfocus="javascript:isToPropagate=true;"
								onKeyDown="textCounter(document.strFormm.Purpose,document.strFormm.remLen3,500)"	onKeyUp="textCounter(document.strFormm.Purpose,document.strFormm.remLen3,500)"></textarea>
                    &nbsp;<font color="red">&nbsp;Max(500 Char) </font>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                  </td>

                  <td  align="Right"><font ><b>Special
                    Instructions&nbsp;:<font color="red">*</font></b></font></td>
							<td  colspan="2">
								<textarea class="formstylemedium" name="SpecialInstructions" rows="4" cols="30"  onfocus="javascript:isToPropagate=true;"
								onKeyDown="textCounter(document.strFormm.SpecialInstructions,document.strFormm.remLen3,500)"	onKeyUp="textCounter(document.strFormm.SpecialInstructions,document.strFormm.remLen3,500)"></textarea>
                    <font color="red">&nbsp;Max(500 Char)</font>&nbsp;
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                  </td>
						</tr>
						<tr height=25>

                  <td  align="Right"><font ><b>Approx
                    Unit Cost&nbsp;: </b></font></td>
							<td  colspan="2">
								<input class="formstylemedium" type="text" name="ApproxUnitCost" size="10" maxlength="9" value="" onfocus="javascript:isToPropagate=true;">&nbsp;
							</td>
							<td  align="Right"><font ><b>Possible Source&nbsp;:&nbsp;</b></font></td>
							<td  colspan="2">
								<input class="formstylemedium" type="text" name="PossibleSource" maxlength="50" size="40" onfocus="javascript:isToPropagate=true;" >&nbsp;
                  </td>
						</tr>
						<tr align="center" valign="middle">
							<td colspan="6"  align="center">
								<input class="formbutton" type="button" name="AddButton" value="Add" style="border: 1 solid; width:50px;"  onclick="javascript:addItemToPurchaseRequest()">&nbsp;
								&nbsp;&nbsp;
								<input class="formbutton" type="button" name="ResetButton" value="Reset" style="border: 1 solid; width:50px; " onclick="javascript:resetItemDetails()">&nbsp;
							</td>
						</tr>
					</table>
					<input type="hidden" name="EditAction" value="">
				</form>
			</td>
		</tr>
		<tr>
			<td>
				<form name="ItemForm">
					<table id="PurchaseRequest" width="98%" align="center" valign="top" cellspacing="2" cellpadding="2" border="0">
						<tr height="25" >
							<td colspan="10" align="center">
								<font ><b>Purchase Request</b></font>
							</td>
						</tr>
						<tr height="25" >
							<td align="center">
								<font ><b>Sl. No.</b></font>
							</td>
							<td align="center">
								<font ><b>Item Description</b></font>
							</td>
							<td align="center">
								<font ><b>Project</b></font>
							</td>
							<td align="center">
								<font ><b>Purpose</b></font>
							</td>
							<td align="center">
								<font ><b>Quantity Required</b></font>
							</td>
							<td align="center">
								<font ><b>Request Type</b></font>
							</td>
							<td align="center">
								<font ><b>Required Date</b></font>
							</td>
							<td align="center">
								<font ><b>Approx Unit Cost</b></font>
							</td>
							<td align="center">
								<font ><b>Possible Source</b></font>
							</td>
							<td align="center">
								<font ><b>Special Instructions</b></font>
							</td>
						</tr>
						<tr>
							<td class="bluelight" colspan="10" align="center">
								<input type="button" class="formbutton" value="Edit" style="border: 1 solid; width:50px;" onclick="javascript:editPurchaseRequestItem();">
								&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="button" class="formbutton" value="Delete" style="border: 1 solid; width:50px; " onclick="javascript:deleteItemFromPurchaseRequest();">
								&nbsp;&nbsp;&nbsp;&nbsp;
								<input type="button" class="formbutton" value="Submit" style="border: 1 solid; width:50px;" onclick="javascript:validateOnSubmit();">
							</td>
						</tr>
						<input name='CheckItem' type='checkbox' value="" style='visibility:hidden'>
						<input name='CheckItem' type='checkbox' value="" style='visibility:hidden'>
						<input name='ProjectId' type='hidden' value="" style='visibility:hidden'>
						<input name='ServiceType' type='hidden' value="" style='visibility:hidden'>
						<input name='Currency' type='hidden' value="" style='visibility:hidden'>

					</table>
				</form>
				<form name="FinalForm">
					<input type="hidden" name="ItemList" value="">
				</form>
			</td>
		</tr>
	</table>
	<script language="javascript">
		var isToPropagate=true;
		var index=2;
		var editIndex=-1;
		function validateProject(ctrl)
		{
			if(isToPropagate==true)
			{
				if(ctrl.value=="0")
				{
					alert("Project Name is mandatory");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateItemDescription(ctrl)
		{
			var regexp = new RegExp (/[0-9a-zA-Z]/);
			ctrl.value=trim(ctrl.value.toUpperCase());
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Item Description is mandatory");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter Alpha-Numeric characters only.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
				{
					var tableReference=document.getElementById('PurchaseRequest');
					var rowCount=tableReference.rows.length;
					--rowCount;
					var arrItems=new Array(rowCount-2);
					for(var i=2;i<rowCount;i++)
						if(document.ItemForm.ProjectId[i-2].value==document.strFormm.Project.value && tableReference.rows[i].cells[1].innerText==document.strFormm.ItemDescription.value)
						{
							alert("For a particular project Items must be unique.");
							isToPropagate=false;
							ctrl.focus();
							return false;
						}
					isToPropagate=true;
				}
			}
			return true;
		}
		function validatePurpose(ctrl)
		{
			var regexp = new RegExp (/[0-9a-zA-Z]/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Purpose is mandatory");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter Alpha-Numeric characters only.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateQuantityRequired(ctrl)
		{
			var regexp = new RegExp (/^[1-9]\d*$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Quantity is mandatory");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter a valid Quantity Required.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}
		function validateRequiredDate(ctrl)
		{
			var regexp=new RegExp(/(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Required Date is mandatory field");
					isToPropagate=false;
					ctrl.focus();
				//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
					return false;
				}
				else if(!regexp.test(ChangeToMMDDYYYY(ctrl.value)))
				{
					alert("Please enter a valid Required Date.");
					isToPropagate=false;
					ctrl.focus();
				//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
					return false;
				}
				else if(isGreaterDate(ChangeToMMDDYYYY(document.strFormm.RequiredDate.value),ChangeToMMDDYYYY('<%=SetDateFormat(Formatdatetime(now(),2))%>')))
				{
					alert ("Required date should be Greater than or equal to Current Date")
					isToPropagate=false;
					ctrl.focus();
				//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
					return false;
				}
				else
					isToPropagate=true;
			}
				return true;
		}
		function validateApproxUnitCost(ctrl)
		{
			var regexp = new RegExp (/^\d+(\.\d\d)?$/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(ctrl.value!=="")
			{
				if(isToPropagate==true)
				{
					if(!regexp.test(ctrl.value))
					{
						alert("Please enter Valid Cost.");
						isToPropagate=false;
						ctrl.focus();
						return false;
					}
					else
						isToPropagate=true;
				}
			}
			return true;
		}
		function validatePossibleSource(ctrl)
		{
			var regexp = new RegExp (/[0-9a-zA-Z]/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(ctrl.value!="")
			{
				if(isToPropagate==true)
				{
					if(!regexp.test(ctrl.value))
					{
						alert("Please enter Alpha-Numeric characters only.");
						isToPropagate=false;
						ctrl.focus();
						return false;
					}
					else
						isToPropagate=true;
				}
			}
			return true;
		}
		function validateSpecialInstructions(ctrl)
		{
			var regexp = new RegExp (/[0-9a-zA-Z]/);
			ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
			if(isToPropagate==true)
			{
				if(ctrl.value=="")
				{
					alert("Special Instructions is mandatory");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else if(!regexp.test(ctrl.value))
				{
					alert("Please enter Alpha-Numeric characters only.");
					isToPropagate=false;
					ctrl.focus();
					return false;
				}
				else
					isToPropagate=true;
			}
			return true;
		}

		function validateOnAdd()
		{
			var referenceProject=document.getElementById("Project");
			var referenceItemDescription=document.getElementById("ItemDescription");
			var referencePurpose=document.getElementById("Purpose");
			var referenceRequiredDate=document.getElementById("RequiredDate");
			var referenceQuantityRequired=document.getElementById("QuantityRequired");
			var referenceApproxUnitCost=document.getElementById("ApproxUnitCost");
			var referencePossibleSource=document.getElementById("PossibleSource");
			var referenceSpecialInstructions=document.getElementById("SpecialInstructions");

			var returnValueProject=validateProject(referenceProject);
			var returnValueItemDescription=validateItemDescription(referenceItemDescription);
			var returnValuePurpose=validatePurpose(referencePurpose);
			var returnValueQuantityRequired=validateQuantityRequired(referenceQuantityRequired);
			var returnRequiredDate=validateRequiredDate(referenceRequiredDate);
			var returnValueApproxUnitCost=validateApproxUnitCost(referenceApproxUnitCost);
			var returnValuePossibleSource=validatePossibleSource(referencePossibleSource);
			var returnValueSpecialInstructions=validateSpecialInstructions(referenceSpecialInstructions);

			if(returnValueProject==true && returnValueItemDescription==true && returnValuePurpose==true && returnValueQuantityRequired==true && returnRequiredDate==true && returnValueApproxUnitCost==true && returnValuePossibleSource==true && returnValueSpecialInstructions==true)
			{
				return true;
			}
			else
				return false;
		}
		function validateOnSubmit()
		{
			if(document.ItemForm.CheckItem.length==2)
			{
				alert("Purchase Request is empty");
				return false;
			}
			else
			{
				var tableReference=document.getElementById('PurchaseRequest');
				var rowCount=tableReference.rows.length;
				--rowCount;
				var arrItems=new Array(rowCount-2);

				for(var i=2;i<rowCount;i++)
				{
					arrItems[i-2]=new Array();
					for(var j=0;j<=11;j++)
					{
						if(j==0)
							arrItems[i-2][j]=document.ItemForm.ProjectId[i-2].value;
						else if(j==10)
							arrItems[i-2][j]=document.ItemForm.ServiceType[i-2].value
						else if(j==11)
							arrItems[i-2][j]=document.ItemForm.Currency[i-2].value
						else
						{
							var temp=tableReference.rows[i].cells[j].innerText;
							while(temp.search(",")!=-1)
								temp=temp.replace(",","&#44;");
							arrItems[i-2][j]=temp;
						}
					}
				}
				document.FinalForm.ItemList.value=arrItems;
				document.FinalForm.method="post";
				document.FinalForm.action="AddPurchaseRequest.asp";
				document.FinalForm.submit();
				return true;
			}
		}
		function addItemToPurchaseRequest()
		{
			if(document.getElementById("AddButton").value=="Add")
			{
				if(validateOnAdd()==true)
				{
					document.getElementById('PurchaseRequest').insertRow(index);
					for(var i=0;i<10;i++)
					{
						document.getElementById('PurchaseRequest').rows[index].insertCell();
					}
					document.getElementById('PurchaseRequest').rows[index].cells[0].innerHTML="<input name='CheckItem' type='checkbox' value='" + index + "'></input><input name='ProjectId' type='hidden' value='" + document.strFormm.Project.value + "'></input><input name='ServiceType' type='hidden' value=''></input></input><input name='Currency' type='hidden' value=''></input>&nbsp;"+(document.getElementById('PurchaseRequest').rows[index].rowIndex-1);
					document.getElementById('PurchaseRequest').rows[index].cells[0].align="center";

					if(document.strFormm.PurchaseOrService[0].checked==true)
						document.ItemForm.ServiceType[index-2].value='0';
					else
						document.ItemForm.ServiceType[index-2].value='1';

					if(document.strFormm.RupeeOrDollar[0].checked==true)
						document.ItemForm.Currency[index-2].value='0';
					else
						document.ItemForm.Currency[index-2].value='1';

					document.getElementById('PurchaseRequest').rows[index].cells[0].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[1].innerText=document.strFormm.ItemDescription.value;
					document.getElementById('PurchaseRequest').rows[index].cells[1].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[1].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[2].innerText=document.strFormm.Project.options[document.strFormm.Project.selectedIndex].text;
					document.getElementById('PurchaseRequest').rows[index].cells[2].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[2].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[3].innerText=document.strFormm.Purpose.value;
					document.getElementById('PurchaseRequest').rows[index].cells[3].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[3].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[4].innerText=document.strFormm.QuantityRequired.value;
					document.getElementById('PurchaseRequest').rows[index].cells[4].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[4].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[5].innerText=(document.strFormm.PurchaseOrService[0].checked==true)?document.strFormm.PurchaseOrService[0].text:document.strFormm.PurchaseOrService[1].text;
					document.getElementById('PurchaseRequest').rows[index].cells[5].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[5].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[6].innerText=document.strFormm.RequiredDate.value;
					document.getElementById('PurchaseRequest').rows[index].cells[6].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[6].className="bluelight";
				//	document.getElementById('PurchaseRequest').rows[index].cells[7].innerText=(document.strFormm.ApproxUnitCost.value!="")?((document.strFormm.RupeeOrDollar[0].checked==true)?(document.strFormm.RupeeOrDollar[0].value+" "+document.strFormm.ApproxUnitCost.value):(document.strFormm.RupeeOrDollar[1].value+" "+document.strFormm.ApproxUnitCost.value)):"-";
					document.getElementById('PurchaseRequest').rows[index].cells[7].innerText=(document.strFormm.ApproxUnitCost.value!="")?((document.strFormm.RupeeOrDollar[0].checked==true)?(document.strFormm.ApproxUnitCost.value+" "+document.strFormm.RupeeOrDollar[0].value):(document.strFormm.ApproxUnitCost.value+" "+document.strFormm.RupeeOrDollar[1].value)):"-";
					document.getElementById('PurchaseRequest').rows[index].cells[7].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[7].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[8].innerText=(document.strFormm.PossibleSource.value=="")?"-":document.strFormm.PossibleSource.value;
					document.getElementById('PurchaseRequest').rows[index].cells[8].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[8].className="bluelight";
					document.getElementById('PurchaseRequest').rows[index].cells[9].innerText=(document.strFormm.SpecialInstructions.value=="")?"-":document.strFormm.SpecialInstructions.value;
					document.getElementById('PurchaseRequest').rows[index].cells[9].align="center";
					document.getElementById('PurchaseRequest').rows[index].cells[9].className="bluelight";

					var cell =  document.getElementById('PurchaseRequest').rows[index].cells[1].innerText + " " + document.getElementById('PurchaseRequest').rows[index].cells[2].innerText ;
					index++;
					resetItemDetails();
				}
				else
					return false;
			}
			else
			{
			//	if(validateOnAdd()==true)
			//	{
					document.getElementById('PurchaseRequest').rows[editIndex].cells[1].innerText=document.strFormm.ItemDescription.value;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[2].innerText=document.strFormm.Project.options[document.strFormm.Project.selectedIndex].text;
					document.ItemForm.ProjectId[editIndex-2].value=document.strFormm.Project.value;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[3].innerText=document.strFormm.Purpose.value;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[4].innerText=document.strFormm.QuantityRequired.value;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[5].innerText=(document.strFormm.PurchaseOrService[0].checked==true)?document.strFormm.PurchaseOrService[0].text:document.strFormm.PurchaseOrService[1].text;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[6].innerText=document.strFormm.RequiredDate.value;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[7].innerText=(document.strFormm.ApproxUnitCost.value!="")?((document.strFormm.RupeeOrDollar[0].checked==true)?(document.strFormm.ApproxUnitCost.value+" "+document.strFormm.RupeeOrDollar[0].value):(document.strFormm.ApproxUnitCost.value+" "+document.strFormm.RupeeOrDollar[1].value)):"-";
					document.getElementById('PurchaseRequest').rows[editIndex].cells[8].innerText=(document.strFormm.PossibleSource.value=="")?"-":document.strFormm.PossibleSource.value;
					document.getElementById('PurchaseRequest').rows[editIndex].cells[9].innerText=(document.strFormm.SpecialInstructions.value=="")?"-":document.strFormm.SpecialInstructions.value;
					if(document.strFormm.PurchaseOrService[0].checked==true)
						document.ItemForm.ServiceType[editIndex-2].value='0';
					else
						document.ItemForm.ServiceType[editIndex-2].value='1';

					if(document.strFormm.RupeeOrDollar[0].checked==true)
						document.ItemForm.Currency[editIndex-2].value='0';
					else
						document.ItemForm.Currency[editIndex-2].value='1';

					editIndex=-1;
					document.getElementById('AddButton').value="Add";
					document.getElementById('ResetButton').disabled=false;
					resetItemDetails();
					return true;
			//	}
			 }
		}
		function resetItemDetails()
		{
			document.strFormm.ItemDescription.value="";
			document.strFormm.Project.value="0";
			document.strFormm.Purpose.value="";
			document.strFormm.QuantityRequired.value="";
			document.strFormm.PurchaseOrService[0].checked=true;
			document.strFormm.RequiredDate.value="";
			document.strFormm.ApproxUnitCost.value="";
			document.strFormm.RupeeOrDollar[0].checked=true;
			document.strFormm.PossibleSource.value="";
			document.strFormm.SpecialInstructions.value="";
		}
		function deleteItemFromPurchaseRequest()
		{
			var selectedItemFlag=false;
			var selectedItemIndex;
			if(document.ItemForm.CheckItem.length==2)
			{
				alert("Purchase Request is empty");
				return false;
			}
			else
			{
				var checkedCount=0;
				for(var i=0;i<document.ItemForm.CheckItem.length-2;i++)
					if(document.ItemForm.CheckItem[i].checked==true)
					{
						if(selectedItemFlag==false)
						{
							selectedItemIndex=i;
							selectedItemFlag=true;
						}
						checkedCount++;
					}
				if(checkedCount==0)
				{
					alert("Please select an item to edit.");
					return false;
				}
				else
				{
					var tableReference=document.getElementById('PurchaseRequest');
					var rowCount;
					var reOrder=false;
					for(var i=document.ItemForm.CheckItem.length-3;i>=0;i--)
					{
						if(document.ItemForm.CheckItem[i].checked==true)
						{
							tableReference.deleteRow(document.ItemForm.CheckItem[i].value);
							--index;
							reOrder=true;
						}
					}
					if(reOrder==true)
					{
						rowCount=tableReference.rows.length;
						--rowCount;
						for(var i=2;i<rowCount;i++)
							document.getElementById('PurchaseRequest').rows[i].cells[0].innerHTML="<input name='CheckItem' type='checkbox' value='" + i + "'></input><input name='ProjectId' type='hidden' value='" + document.ItemForm.ProjectId[i-2].value + "'></input><input name='ServiceType' type='hidden' value='" + document.ItemForm.ServiceType[i-2].value + "'></input></input><input name='Currency' type='hidden' value='" + document.ItemForm.Currency[i-2].value +"'></input>&nbsp;"+(document.getElementById('PurchaseRequest').rows[i].rowIndex-1);
					}
					return true;
				}
			}
		}
		function editPurchaseRequestItem()
		{
			var selectedItemFlag=false;
			var selectedItemIndex;
			if(document.ItemForm.CheckItem.length==2)
			{
				alert("Purchase Request is empty");
				return false;
			}
			else
			{
				var checkedCount=0;
				for(var i=0;i<document.ItemForm.CheckItem.length-2;i++)
					if(document.ItemForm.CheckItem[i].checked==true)
					{
						if(selectedItemFlag==false)
						{
							selectedItemIndex=i;
							selectedItemFlag=true;
						}
						checkedCount++;
					}
				if(checkedCount==0)
				{
					alert("Please select an item to edit.");
					return false;
				}
				else if(checkedCount>1)
				{
					alert("Please select only one item to edit.");
					return false;
				}
				else
				{
					var referenceTableRow=document.getElementById('PurchaseRequest').rows[document.ItemForm.CheckItem[selectedItemIndex].value];
					document.strFormm.ItemDescription.value=referenceTableRow.cells[1].innerText;
					document.strFormm.Project.value=document.ItemForm.ProjectId[selectedItemIndex].value;
					document.strFormm.Purpose.value=referenceTableRow.cells[3].innerText;
					document.strFormm.QuantityRequired.value=referenceTableRow.cells[4].innerText;

					if(referenceTableRow.cells[5].innerText=="Purchase")
						document.strFormm.PurchaseOrService[0].checked=true;
					else
						document.strFormm.PurchaseOrService[1].checked=true;

					document.strFormm.RequiredDate.value=referenceTableRow.cells[6].innerText;

					if(referenceTableRow.cells[7].innerText=="-")
						document.strFormm.ApproxUnitCost.value="";
					else
					{
						var arrTemp=referenceTableRow.cells[7].innerText.split(" ")
						document.strFormm.ApproxUnitCost.value=arrTemp[0];
						if(arrTemp[1]=="Rupee(s)")
							document.strFormm.RupeeOrDollar[0].checked=true;
						else
							document.strFormm.RupeeOrDollar[1].checked=true;
					}

					if(referenceTableRow.cells[8].innerText=="-")
						document.strFormm.PossibleSource.value="";
					else
						document.strFormm.PossibleSource.value=referenceTableRow.cells[8].innerText;

					if(referenceTableRow.cells[9].innerText=="-")
						document.strFormm.SpecialInstructions.value="";
					else
						document.strFormm.SpecialInstructions.value=referenceTableRow.cells[9].innerText;

					editIndex=document.ItemForm.CheckItem[selectedItemIndex].value;
					document.ItemForm.CheckItem[selectedItemIndex].checked=false;
					document.getElementById('AddButton').value="Save"
					document.getElementById('ResetButton').disabled=true;
					return true;
				}
				return true;
			}
		}
		function callOpenCalendar(ctrl)
		{
			if(ctrl.value=="")
				openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','RequiredDate',150,300);
		}
		function ChangeToMMDDYYYY(strDate)
		{
			var datarr;
			var strmon;
			datarr=strDate.split("-")
			switch(datarr[1])
			{
				case 'Jan':
					strmon="01";
					break;
				case 'Feb':
					strmon="02";
					break;
				case 'Mar':
					strmon="03";
					break;
				case 'Apr':
					strmon="04";
					break;
				case 'May':
					strmon="05";
					break;
				case 'Jun':
					strmon="06";
					break;
				case 'Jul':
					strmon="07";
					break;
				case 'Aug':
					strmon="08";
					break;
				case 'Sep':
					strmon="09";
					break;
				case 'Oct':
					strmon="10";
					break;
				case 'Nov':
					strmon="11";
					break;
				case 'Dec':
					strmon="12";
					break;
			}
			if(datarr[0]=="1"||datarr[0]=="2"||datarr[0]=="3"||datarr[0]=="4"||datarr[0]=="5"||datarr[0]=="6"||datarr[0]=="7"||datarr[0]=="8"||datarr[0]=="9")
				datarr[0]="0"+datarr[0];

			return (strmon+'/'+datarr[0]+'/'+datarr[2]);
		}
		function ChangeToVelankaniFormatDate(strDate)
		{
			var datarr;
			var strmon;
			datarr=strDate.split("/")
			switch(datarr[0])
			{
				case "1":
					strmon="Jan";
					break;
				case "2":
					strmon="Feb";
					break;
				case "3":
					strmon="Mar";
					break;
				case "4":
					strmon="Apr";
					break;
				case "5":
					strmon="May";
					break;
				case "6":
					strmon="Jun";
					break;
				case "7":
					strmon="Jul";
					break;
				case "8":
					strmon="Aug";
					break;
				case "9":
					strmon="Sep";
					break;
				case "10":
					strmon="Oct";
					break;
				case "11":
					strmon="Nov";
					break;
				case "12":
					strmon="Dec";
					break;
			}
			return (datarr[1] + "-" + strmon + "-" + datarr[2]);
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

<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>
<!--#include file="../includes/main_page_close.asp"-->