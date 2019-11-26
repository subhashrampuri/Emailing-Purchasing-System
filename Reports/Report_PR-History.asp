
<%
	Response.Buffer =True
	Server.ScriptTimeout = 10000
	Response.ContentType = "application/vnd.ms-excel"

	FromDate = Trim(Request.Form("hdFromDate"))
	ToDate = Trim(Request.Form("hdToDate"))
	EmployeeID = Trim(Request.Form("hdEmpID"))
	
	Response.AddHeader "content-disposition","attachment;filename=PR-History_" & FromDate & "_" & ToDate & ".xls"

%>

<!--#include file="../../includes/main_page_header.asp"-->
	<table width="90%" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#999999">
	  <tr bgcolor="#FFFFFF"  valign="middle"> 
		<%
		
			strSql = " Select RequisitionID,dbo.fn_Psystem_PurchaseRequisitionNo(RequisitionNum) as RequisitionNo, " & _ 
				" dbo.fn_TSystem_GetVelankaniFormatDate(RequisitionDate) as RequestedDate from tbl_Psystem_PurchaseRequestMaster " & _ 
				" where EmployeeID = '" & EmployeeID & "' and RequisitionDate between '" & FromDate & "' and '"& ToDate &"' "
			Call RunSql(strSql,rsReq)	
			
			if Not rsReq.EOF then
				While Not rsReq.EOF
			%>
		<td bgcolor="#FFFFFF" colspan="12"> 
		  <div align="center"><b>Purchase Request No : </b><%=rsReq("RequisitionNo")%> 
			&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<b>Requested Date : </b><%=rsReq("RequestedDate")%> 
		  </div>
		</td>
	  </tr>
	  <tr bgcolor="#FFFFFF" valign="middle"> 
		<td> 
		  <div align="center"><b>Sl.No</b></div>
		</td>
		<td > 
		  <div align="center" ><b>Item Description</b></div>
		</td>
		<td> 
		  <div align="center"><b>Project</b></div>
		</td>
		<td> 
		  <div align="center"><b>Purpose</b></div>
		</td>
		<td> 
		  <div align="center"><b>Quantity Requested</b></div>
		</td>
		<td> 
		  <div align="center"><b>Quantity Approved</b></div>
		</td>
		<td> 
		  <div align="center"><b> Approx Unit Cost</b></div>
		</td>
		<td> 
		  <div align="center"><b>Special Instructions</b></div>
		</td>
		<td colspan="4"> 
		  <div align="center"><b>Status</b></div>
		</td>
	  </tr>
	  <%
			MySql = " Select a.ItemDescription,dbo.fn_TimeSheet_GetProjectName(a.ProjectId) as Project,a.Purpose, " & _ 
				" a.QuantityRequested,b.QuantityApproved,a.ApproxUnitCost,dbo.fn_PSystem_isRupeeOrDollar(a.RupeeOrDollar) as Currency, " & _ 
				" a.SpecialInstruction,b.Status from tbl_Psystem_PurchaseRequestTransaction a  " & _ 
				" left outer join tbl_Psystem_TransactionDetails b on a.RequisitionID = b.RequisitionID and a.ProjectID = b.ProjectID " & _ 
				" where a.RequisitionID  = "& rsReq("RequisitionID") &" "

			Call RunSql(MySql,objRs)	
			if NOT objRs.EOF then
				i = 1
				While NOT objRs.EOF 
					
				if objRs("QuantityApproved") <> "" then
					QtyApproved = objRs("QuantityApproved")
				else
					QtyApproved = 0
				end if

				if objRs("Status") = 1 then
					sStatus = "Approved"
				elseif objRs("Status") = 2 then
					sStatus = "Partially Approved"				
				elseif objRs("Status") = 3 then
					sStatus = "Rejected"
				elseif objRs("Status") = 4 then
					sStatus = "Approved and In Process"
				else
					sStatus = "New Request"									
				end if
					
			%>
	  <tr bgcolor="#FFFFFF"> 
		<td> 
		  <div align="center"><%=i%></div>
		</td>
		<td nowrap><%=objRs("ItemDescription")%></td>
		<td><%=objRs("Project")%></td>
		<td><%=objRs("Purpose")%></td>
		<td> 
		  <div align="center"><%=objRs("QuantityRequested")%></div>
		</td>
		<td> 
		  <div align="center"><%=QtyApproved%></div>
		</td>
		<td><%=objRs("ApproxUnitCost") & " " & objRs("Currency") %></td>
		<td><%=objRs("SpecialInstruction")%></td>
		<td colspan="4"> 
		  <div align="center"><%=sStatus%></div>
		</td>
	  </tr>
	  <%
			i = i + 1
			objRs.MoveNext
			Wend
			End if
			objRs.Close
		%>
	  <tr bgcolor="#FFFFFF"> 
		<td colspan="12">&nbsp;</td>
	  </tr>
	  <%
			rsReq.movenext
			wend
		else
			Response. Write "<tr><td align='center' bgcolor='#ffffff' colspan='11'><b>No Records Found</b></td></tr>"
			end if
			set rsReq = NOTHING
	  %>
	</table>
<br>
<!--#include file="../../includes/connection_close.asp"-->