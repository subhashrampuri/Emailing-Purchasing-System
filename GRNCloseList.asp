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
  function Close_GRN()
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
		if (document.getElementById("ChkAction_"+k).checked == false)
		  {
			cnt=cnt+1;
		  }
	} //for loop

	  if (half==cnt)
	  {
	   alert('Select GRN to Close');
	   return false;
	  }

 	document.strFormm.method="post";
	document.strFormm.action="GRN_Close.asp";
	document.strFormm.submit();

  }
</Script>
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

 %>
      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr height="25" class="blue" align="center">
          <td align="center" width="95%"> <font color=#ffffff><b>CLOSE GOODS RECEIVED NOTE (GRN)</b></font></td>
          <td align="right" width="5%"><p style="margin-right:10"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></p></td>
        </tr>
        <tr hight="25">
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr height="25" align="left">
          <td colspan="2" > </td>
        </tr>
        <tr>
          <td align="center" >&nbsp; </td>
        </tr>
        <tr>
          <td colspan="2">
            <form name="strFormm">
              <table cellspaning="2" cellpadding="2" border="0" width="50%" align="center">
                <tr class="blue" align="center">
                  <td align="Center"><font color="#ffffff"><b>Sl.No</b></font></td>
                  <td align="center"><font color="#ffffff"><b>GOODS RECEIVED NOTE
                    No</b></font></td>
                  <td align="center"><font color="#ffffff"><b>Close </b></font></td>
                </tr>
                <%
					Dim i

					sql ="Select distinct GRNNo,RequisitionId,PurOrderNo from tbl_Psystem_GRN where isAccepted = 1 and isGRNClosed = 0"
					call RunSql(sql,rsItems)

					i = 1
					if rsItems.Eof = false then
					while Not rsItems.EOF
					GRNNo =	rsItems("GRNNo")
					sql = " Select GRNNum  from tbl_Psystem_GRN where GRNNo = "& GRNNo &" "
					call RunSql(sql,rsGRNNum)
					if rsGRNNum.Eof = false then
						GRNNum = rsGRNNum("GRNNum")
					end if
					rsGRNNum.Close

				%>
                <tr bgcolor="<%=gsBGColorLight%>">
                  <td align="center"><%=i%></td>
                  <td align="center"><font color="black"><%=GetGRNNo(GRNNum)%></font></td>
                  <td align="center">
                    <input type="checkbox" name="ChkAction_<%=i%>" value="<%=i%>">
                  </td>
                  <input type="hidden" name="hdReqId_<%=i%>" value="<%=rsItems("RequisitionId")%>">
                  <input type="hidden" name="hdGRNNo_<%=i%>" value="<%=rsItems("GRNNo")%>">
                  <input type="hidden" name="hdPurOrdNo_<%=i%>" value="<%=rsItems("PurOrderNo")%>">
                </tr>
                <%
					i = i + 1
					rsItems.movenext
					Wend
					else
						Response.Write "<tr ><td align='center' colspan='3'><br><b><font color='red'> No records found </font></b></td></tr>"
					end if
					rsItems.Close
					'iCount=rsItems.recordcount


					sql= "Select  Count(distinct GRNNo) as iCount from tbl_Psystem_GRN where isAccepted = 1 and isGRNClosed = 0"
					Call RunSql(sql,rsCnt)
					iCount = rsCnt("iCount")
					'response.write iCount

				%>
                <tr >
                  <td align="center" bgcolor="<%=gsBGColorLight%>" colspan="2">&nbsp;</td>
                  <td align="center" bgcolor="<%=gsBGColorLight%>">
                    <input type="button" class=formbutton name="Close" value="Close" style="border: 1 solid" onclick="Close_GRN();">
                  <input type="hidden" name="hdCount" value="<%=iCount%>">
                  </td>
                </tr>
                <tr class="blue">
                  <td align="center" colspan="3">&nbsp;</td>
                </tr>
              </Table>
            </form>
          </td>
        </tr>
      </table>
<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->