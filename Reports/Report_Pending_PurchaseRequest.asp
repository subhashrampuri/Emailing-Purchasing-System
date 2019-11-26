
<%
	Response.Buffer =True
	Server.ScriptTimeout = 10000
	Response.ContentType = "application/vnd.ms-excel"

	'Response.Charset = "GB2312"
	'Response.Codepage = "936"



%>

<!--#include file="../../includes/main_page_header.asp"-->


<table width="100%" border="1" cellspacing="2" cellpadding="2">
  <tr bgcolor="#CCCCCC" valign="middle">
    <td height="30">
      <div align="center"><b>Sl.No</b></div>
    </td>
    <td height="30" >
      <div align="center" ><b>Requisition No</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Item Desctiption</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Project</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Quantity Requested</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Possible Source</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Special Instructions</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Approx Unit Cost</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Service Type</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Requested Date</b></div>
    </td>
    <td height="30">
      <div align="center"><b>Requested Employee</b></div>
    </td>
  </tr>
  <%
	FromDate = Trim(Request.Form("hdFromDate"))
	ToDate = Trim(Request.Form("hdToDate"))

	sql = " Select dbo.fn_Psystem_PurchaseRequisitionNo(b.RequisitionNum) as RequisitionID,a.ItemDescription,dbo.fn_TimeSheet_GetProjectName(a.ProjectId) as Project, " & _
		" a.QuantityRequested,a.QuantityApproved,(a.QuantityRequested - a.QuantityApproved) as QtyPending,a.PossibleSource,a.SpecialInstruction, " & _
		" a.ApproxUnitCost,dbo.fn_PSystem_isRupeeOrDollar(a.RupeeOrDollar) as Currency,dbo.fn_PSystem_isPurchaseOrService(a.PurchaseOrService) as ServiceType, " & _
		" dbo.fn_TSystem_GetVelankaniFormatDate(b.RequisitionDate) as RequestedDate,(dbo.fn_TSystem_EmployeeName(b.EmployeeID)+' ('+ b.EmployeeID+')') as Employee " & _
		" from tbl_Psystem_PurchaseRequestTransaction a,tbl_Psystem_PurchaseRequestMaster b where a.RequisitionId = b.RequisitionId and a.Status = 2 " & _
		" and b.RequisitionDate between '" & FromDate &"' and '" & ToDate & "' "
		Call RunSql(sql,objRs)

	Response.AddHeader "content-disposition","attachment;filename=Purchase_Request_" & FromDate & "_" & ToDate & ".xls"

		i = 1
		If Not objRs.EOF Then
			while Not objRs.EOF
	%>

  <tr>
    <td>
      <div align="center"><%=i%></div>
    </td>
    <td><%=objRs("RequisitionID")%></td>
    <td><%=objRs("ItemDescription")%></td>
    <td><%=objRs("Project")%></td>
    <td>
      <div align="center"><%=objRs("QuantityRequested")%></div>
    </td>
    <td><%=objRs("PossibleSource")%></td>
    <td><%=objRs("SpecialInstruction")%></td>
    <td>
      <div align="center"><%=objRs("ApproxUnitCost") & " " & objRs("Currency")%></div>
    </td>
    <td><%=objRs("ServiceType")%></td>
    <td>
      <div align="center"><%=objRs("RequestedDate")%></div>
    </td>
    <td><%=objRs("Employee")%></td>
  </tr>

  <%
	i = i + 1
	  objRs.movenext
	  wend
	 else
		Response.Write "No Records Found"
		end if
		set objRs = NOTHING
  %>
</table>


<!--#include file="../../includes/connection_close.asp"-->