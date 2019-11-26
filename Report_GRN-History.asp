
<%

	Server.ScriptTimeout = 10000
	Response.ContentType = "application/vnd.ms-excel"

	Response.Charset = "GB2312"
	Response.Codepage = "936"

	PurOrdNo = Trim(Request.Form("hdPurOrdNo"))
	iGRNNo = Trim(Request.Form("hdGRNNo"))

%>

<!--#include file="../../includes/main_page_header.asp"-->

	<%
		  sql = " Select dbo.fn_Psystem_GRNNo(GRNNo) as GRNNo,dbo.fn_Psystem_PurchaseRequisitionNo(RequisitionId) as RequisitionNo, " & _
				" dbo.fn_Psystem_PurchaseOrderNo(PurOrderNo) as PurOrderNo,PartyChallanNo,dbo.fn_TSystem_GetVelankaniFormatDate(PartyChallanDate) as PartyChallanDate, " & _
				" SecurityEntryNo,dbo.fn_TSystem_GetVelankaniFormatDate(DeliveryDate) as DeliveryDate,LLRRNo,VehicleNo,SupplierName, " & _
				" dbo.fn_Psystem_GetSupplierAddress(SupplierName) as SupplierAddress from tbl_Psystem_GRN  " & _
				" where PurOrderNo = " & PurOrdNo & " and GRNNo = "& iGRNNo &" "
		  call RunSql(sql,objRs)

		  if objRs.EOF = false then
		  	'while Not objRs.EOF
		  %>
            <table width="90%" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#999999">
              <tr bgcolor="#FFFFFF">
                <td>
                  <p style="margin-left=10"><b>GRN No :</b></p>
                </td>
                <td nowrap>
                  <p style="margin-left=10"><%=objRs("GRNNo")%></p>
                </td>
                <td nowrap>
                  <p style="margin-left=10"><b>Purchase Order No :</b></p>
                </td>
                <td>
                  <p style="margin-left=10"><%=objRs("PurOrderNo")%></p>
                </td>
                <td>
                  <p style="margin-left=10"><b>Requisition No :</b></p>
                </td>
                <td colspan="2">
                  <p style="margin-left=10"><%=objRs("RequisitionNo")%></p>
                </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td valign="top" nowrap>
                  <p style="margin-left=10"><b>Supplier Name : </b></p>
                </td>
                <td valign="top" >
                  <p style="margin-left=10"><%=objRs("SupplierName")%></p>
                </td>
                <td valign="top">
                  <p style="margin-left=10"><b>Supplier Address :</b></p>
                </td>
                <td colspan="4" valign="top">
                  <p style="margin-left=10"><%=objRs("SupplierAddress")%></p>
                </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td>
                  <p style="margin-left=10"><b>Party Challan No :</b></p>
                </td>
                <td>
                  <p style="margin-left=10"><%=objRs("PartyChallanNo")%></p>
                </td>
                <td>
                  <p style="margin-left=10"><b>Party Challan Date :</b></p>
                </td>
                <td>
                  <p style="margin-left=10"><%=objRs("PartyChallanDate")%></p>
                </td>
                <td>
                  <p style="margin-left=10"><b>Security Gate No :</b></p>
                </td>
                <td colspan="2">
                  <p style="margin-left=10"><%=objRs("SecurityEntryNo")%></p>
                </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td>
                  <p style="margin-left=10"><b>Delivery Date :</b></p>
                </td>
                <td>
                  <p style="margin-left=10"><%=objRs("DeliveryDate")%></p>
                </td>
                <td>
                  <p style="margin-left=10"><b>LL/ RR No :</b></p>
                </td>
                <td>
                  <p style="margin-left=10"><%=objRs("LLRRNo")%></p>
                </td>
                <td>
                  <p style="margin-left=10"><b>Vehicle No :</b></p>
                </td>
                <td colspan="2">
                  <p style="margin-left=10"><%=objRs("VehicleNo")%></p>
                </td>
              </tr>
              <tr bgcolor="#FFFFFF">
                <td>
                  <div align="center"><b>Sl No </b></div>
                </td>
                <td>
                  <div align="center"><b>Item Description</b></div>
                </td>
                <td>
                  <div align="center"><b>Quantity Received</b></div>
                </td>
                <td>
                  <div align="center"><b>Quantity Accepted </b></div>
                </td>
                <td>
                  <div align="center"><b>Quantity Rejected</b></div>
                </td>
                <td>
                  <div align="center"><b>Status </b></div>
                </td>
                <td>
                  <div align="center"><b>Remarks</b></div>
                </td>
              </tr>
              <%
			  sql = " select ItemDescription,QtyReceived,QtyAccepted,QtyRejected,isAccepted,RemarksOnAccOrRej " & _
			  	" from tbl_Psystem_GRN where PurOrderNo = " & PurOrdNo & " and GRNNo = "& iGRNNo &" "
				call RunSql(sql,rsGRN)
				i = 1
				if rsGRN.EOF = false then
					While NOT rsGRN.EOF

					if rsGRN("isAccepted") =  1 then
						Status = "Accpeted"
					elseif rsGRN("isAccepted") = 2 then
						Status = "Rejected"
					else
						Status = "Pending"
					end if
			  %>
              <tr bgcolor="#FFFFFF">
                <td>
                  <div align="center"><%=i%></div>
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
                  <div align="center"><%=Status%></div>
                </td>
                <td><%=rsGRN("RemarksOnAccOrRej")%></td>
              </tr>

		    <%
				i = i +1
				rsGRN.MoveNext
				Wend
				end if

			%>
			<%
				rsGRN.Close
				'objRs.MoveNext
				'Wend
				end if
				objRs.Close
			%>
			  <tr bgcolor="#FFFFFF">
                <td colspan="7">&nbsp;</td>
              </tr>
            </table>
	<br>
<!--#include file="../../includes/connection_close.asp"-->