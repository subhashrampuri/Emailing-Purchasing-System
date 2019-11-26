
<%
	Response.Buffer =True
	Server.ScriptTimeout = 10000
	Response.ContentType = "application/vnd.ms-excel"

	Response.Charset = "GB2312"
	Response.Codepage = "936"

	FromDate = Trim(Request.Form("hdFromDate"))
	ToDate = Trim(Request.Form("hdToDate"))
	
	Response.AddHeader "content-disposition","attachment;filename=GRN_" & FromDate & "_" & ToDate & ".xls"

%>

<!--#include file="../../includes/main_page_header.asp"-->
		  <%
		  	'sql = " Select distinct dbo.fn_Psystem_PurchaseOrderNo(PurOrderNo) as PurchaseOrderNo,PurOrderNo,dbo.fn_Psystem_GRNNo(GRNNo) as Goods,GRNNo,Remarks, " & _
			'	" dbo.fn_TSystem_GetVelankaniFormatDate(DeliveryDate) as DeliveryDate from tbl_Psystem_GRN " & _ 
			'	" Where DeliveryDate between '" & FromDate & "' and '" & ToDate & "' Order  by PurOrderNo "
			
			Sql = " Select distinct dbo.fn_Psystem_PurchaseOrderNo(a.PurOrderNum) as PurchaseOrderNo,a.PurOrderNo,b.GRNNo as GRNNo, " & _ 
				" b.Remarks as Remarks,dbo.fn_Psystem_GRNNo(b.GRNNum) as Goods,dbo.fn_TSystem_GetVelankaniFormatDate(b.Deliverydate) as Deliverydate from  " & _ 
				" tbl_Psystem_PurchaseOrder a,tbl_Psystem_GRN b where a.PurOrderNo = b.PurOrderNo and b.DeliveryDate " & _ 
				" between '"& FromDate &"' and '"& ToDate &"' Order  by a.PurOrderNo "
				
				Call Runsql(sql,objRs)
			if objRs.Eof = False then	
				Do While Not objRs.Eof
				
				PurOrdNo = objRs("PurOrderNo")
				DDate = objRs("DeliveryDate")
				Goods = objRs("Goods")
				GRNNo = objRs("GRNNo")
				Remarks = objRs("Remarks")
				
		
				
		  %>
            <table width="90%" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
              <tr bgcolor="#FFFFFF"> 
                <td > 
                  <p style="margin-left:10"><b>GRN No :</b></p>
                </td>
                <td> 
                  <p style="margin-left:10"><%=Goods%></p>
                </td>
                <td> 
                  <p style="margin-left:10"><b>Purchase Order No :</b></p>
                </td>
                <td> 
                  <p style="margin-left:10"><%=objRs("PurchaseOrderNo")%></p>
                </td>
                <td> 
                  <p style="margin-left:10"><b>Date :</b></p>
                </td>
                <td><p style="margin-left:10"><%=DDate%></p></td>
              </tr>
              <tr bgcolor="#FFFFFF" align="center"> 
                <td><b>Sl No</b></td>
                <td><b>Item Description</b></td>
                <td><b>Quantity Received</b></td>
                <td><b>Quantity Accepted</b></td>
                <td colspan="2"><b>Quantity Rejected</b></td>
              </tr>
              <%
				  sql = "Select ItemDescription,QtyReceived,QtyAccepted,QtyRejected from tbl_Psystem_GRN  " & _ 
				  	" where PurOrderNo = "& PurOrdNo &" and GRNNo = "& GRNNo &" "	
				  Call RunSql(sql,rsGRN)
				 ' Response.write sql
				  
				  i = 1
				  if rsGRN.EOF = false then
				  Do While Not rsGRN.EOF
					
			  %>
              <tr bgcolor="#FFFFFF" align="center"> 
                <td><%=i%></td>
                <td><%=rsGRN("ItemDescription")%></td>
                <td><%=rsGRN("QtyReceived")%></td>
                <td><%=rsGRN("QtyAccepted")%></td>
                <td colspan="2"><%=rsGRN("QtyRejected")%></td>
              </tr>
              <%
					i = i + 1
					rsGRN.movenext
					Loop
					end if
					rsGRN.Close
			  %>
              <tr bgcolor="#FFFFFF" > 
                <td> 
                  <p style="margin-left:10"><b>Remarks :</b></p>
                </td>
                <td colspan="5"><p style="margin-left:10"><%=Remarks%></p></td>
              </tr>
              <br>
            </table>
            <br><br>
			<% 
				objRs.movenext
				Loop
				objRs.Close
			else
				Response. Write "<tr><td align='center' bgcolor='#ffffff' Colspan='2'><b>No Records Found</b></td></tr>"
			end if
			%>
	


<br>



<!--#include file="../../includes/connection_close.asp"-->