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
<Script language=JavaScript src="../includes/javascript/validate.js" type=text/javascript></SCRIPT>
 <script language="JavaScript" type="text/javascript">
  function submit_approve()
	{
	var x=document.strFormm.elements.length;
	var p=0;
	var half=0;
	var cnt=0;
	//alert(x);
	for(i=0;i<x;i++)
	{
	if(document.strFormm.elements[i].type=="checkbox")
		{
		p=p+1;

		}
	}

	half=p;

	for(k=1;k<=half;k++)
	{
		if (document.getElementById("action_"+k).checked == false)
		  {
			cnt=cnt+1;
		  }
	} //for loop

	  if (half==cnt)
	  {
	   alert('Select an item to approve');
	   return false;
	  }

	document.strFormm.method="post";
	document.strFormm.action="GRN_Approvers.asp";
	document.strFormm.submit();

  }

  function submit_reject()
	{
	var x=document.strFormm.elements.length;
	var p=0;
	var half=0;
	var cnt=0;
	//alert(x);
	for(i=0;i<x;i++)
	{
	if(document.strFormm.elements[i].type=="checkbox")
		{
		p=p+1;

		}
	}
	half=p;
	for(k=1;k<=half;k++)
	{
		if (document.getElementById("action_"+k).checked == false)
		  {
			cnt=cnt+1;
		  }
	} //for loop

	  if (half==cnt)
	  {
	   alert('Select an item to Reject');
	   return false;
	  }

	document.strFormm.method="post";
	document.strFormm.action="GRN_Reject.asp";
	document.strFormm.submit();
  }

</script>

<%
	ReceivedEditAction="Add"
	EmployeeId=Session("Employee_Id")

	sql="select EmployeeId from tbl_WKI_Admin where EmployeeId='" & EmployeeId  & "'"
	call RunSQL(sql,rsAdmin)
	if not rsAdmin.eof then
		isAdmin=true
	else
		isAdmin=false
	end if
	rsAdmin.close()

	sql="select DeptName from VSPL_Managers A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsDepartment)
	if not rsDepartment.eof then
		isManager=true
		if rsDepartment(0)="HR" then
			isHRManager=true
		elseif rsDepartment(0)="ITIT" then
			isITITManager=true
		else
			isHRManager=false
			isITITManager=false
		end if
	else
		isManager=false
	end if
	rsDepartment.close()

	sql="select DeptName from VSPL_EmployeeHREntry A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsDepartment)
	if not rsDepartment.eof then
		if rsDepartment(0)="HR" then
			isHRMember=true
		elseif rsDepartment(0)="ITIT" then
			isITITMember=true
		else
			isHRMember=false
			isITITMember=false
		end if
	else
		isManager=false
	end if
	rsDepartment.close()


		GoodsNo = Request.form("hdGRNNo")
		'Response.write GoodsNo

%>

      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr height="25" class="blue" align="center">
          <td align="center" colspan="5"> <font color=#ffffff><b>GOODS RECEIVED
            NOTE (GRN)</b></font></td>
        </tr>
        <tr hight="25">
          <td colspan="5">&nbsp;</td>
        </tr>
		<tr height="25" align="left">
          <td colspan="5">&nbsp;</td>
        </tr>
        <tr>
          <td colspan="5" vAlign="top">
            <form name="strFormm">
              <table width="80%" border="0" cellspacing="2" cellpadding="2" align="center">
                <tr class="blue">
                  <td>
                    <div align="center"><font color="#ffffff"><b>Sl.No</b></font></div>
                  </td>
                  <td>
                    <div align="center"><font color="#ffffff"><b>Item Description</b></font></div>
                  </td>
                  <td>
                    <div align="center"><font color="#ffffff"><b>Quantity Received</b></font></div>
                  </td>
                  <td>
                    <div align="center"><font color="#ffffff"><b>Quantity Accepted</b></font></div>
                  </td>
                  <td>
                    <div align="center"><font color="#ffffff"><b>Quantity Rejected</b></font></div>
                  </td>
                  <td>
                    <div align="center"><font color="#ffffff"><b>Remarks for Approve / Reject</b></font></div>
                  </td>
                </tr>
                <%
				sql = "Select * from tbl_Psystem_GRN where GRNNo = "& GoodsNo &" and isAccepted = 0 and isGRNClosed = 0"
				call RunSql(sql,rsGRN)
				i = 1
				if rsGRN.EOF = false then
				while NOT rsGRN.EOF

				%>
                <tr bgcolor="<%=gsBGColorLight%>">
                  <td>
                    <div align="center">
                      <input type="checkbox" name="action_<%=i%>" value="<%=i%>">
                      <%=i%> </div>
                  </td>
                  <td>
                    <div align="center"><%=rsGRN("ItemDescription")%></div>
                  </td>
                  <td>
                    <div align="center"><%=rsGRN("QtyReceived")%></div>
                  </td>
                  <td>
                    <div align="center"><%=rsGRN("QtyAccepted")%></div>
                  </td>
                  <td>
                    <div align="center"><%=rsGRN("QtyRejected")%></div>
                  </td>
                  <td>
                    <div align="center">
                      <input type="text" name="Remarks_<%=i%>" maxlength="255">
                    </div>
					<input type="hidden" name="hdItemDesc_<%=i%>" value="<%=rsGRN("ItemDescription")%>">
					<input type="hidden" name="hdQtyAccepted_<%=i%>" value="<%=rsGRN("QtyAccepted")%>">
					<input type="hidden" name="hdPurOrdNo_<%=i%>" value="<%=rsGRN("PurOrderNo")%>">
					<input type="hidden" name="hdReqNo_<%=i%>" value="<%=rsGRN("RequisitionId")%>">
					<input type="hidden" name="hdGRNNo_<%=i%>" value="<%=rsGRN("GRNNo")%>">
                  </td>
                </tr>
                <%
				i = i  + 1
				rsGRN.movenext
				Wend
				iCount=rsGRN.recordcount
				rsGRN.Close

				else
					Response.redirect ("GRN_List.asp")
				end if
				%>
                <tr>
                  <td colspan="6" align="center">&nbsp;

                  </td>
                </tr>
                <tr>
                  <td colspan="6" align="center">
                    <input class=formbutton type=button value="Approve" name="Approve" style="border: 1 solid" onclick="submit_approve();">
                    &nbsp;&nbsp;&nbsp;
                    <input class=formbutton type=button value="Reject" name="Reject" style="border: 1 solid" onclick="submit_reject();" >
				  	<input type="hidden" name="hdCount" value="<%=iCount%>">
				  </td>
                </tr>
              </table>
              &nbsp;
            </form>
          </td>
        </tr>
      </table>
<p align="center">
<a href="../../imorfusadmin/"><%=dictLanguage("Return_Business_Console")%></a>
</p>
<!--#include file="../includes/main_page_close.asp"-->