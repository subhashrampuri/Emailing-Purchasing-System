
<!--#include file="../includes/MailDesign.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/Connection_open.asp"-->
<!--#include file="../includes/style.asp"-->

<%
	GRNNo = request.querystring("GRNNo")

%>
<title><%=gsSiteName%></title>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="3" align="right"><img src="<%=gsSiteRoot%>gif/logo/Velankani_logo.gif" border = "0" ></td>
  </tr>
  <tr>
    <td align="center" colspan="3"><font size="5"><b>GOODS RECEIVED NOTE (GRN)</b></font></td>
  </tr>
  <tr>
    <td align="center" colspan="3">
      <hr>
    </td>
  </tr>
  <tr>
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td align="Left" colspan="3">
      <table width="95%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor=#666666>
        <%
		  	if GRNNo <> "" then
		  	sql = "Select distinct GRNNo,PurOrderNo,RequisitionId,PartyChallanNo,PartyChallanDate,SecurityEntryNo,DeliveryDate,LLRRNo,vehicleNo,SupplierName,Remarks from tbl_PSystem_GRN where GRNNo = "& GRNNo &" "
			Call RunSql(sql,rsGRN)

			if rsGRN.EOF = false then
				GRNNo = rsGRN("GRNNo")
				sql = " Select GRNNum  from tbl_Psystem_GRN where GRNNo = "& GRNNo &" "
				call RunSql(sql,rsGRNNum)
				if rsGRNNum.Eof = false then
					GRNNum = rsGRNNum("GRNNum")
				end if
				rsGRNNum.Close

				PurOrderNo = rsGRN("PurOrderNo")
				sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrderNo &" "
				call RunSql(sql,rsPONum)
				if rsPONum.Eof = false then
					PurOrderNum = rsPONum("PurOrderNum")
				end if
				rsPONum.Close

				RequisitionId = rsGRN("RequisitionId")
				sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& RequisitionId &" "
				call RunSql(sql,rsRec)
				if rsRec.Eof = false then
					ReqNum = rsRec("RequisitionNum")
				end if
				rsRec.Close
				
				sSupName = rsGRN("SupplierName")
				PartyChallanNo = rsGRN("PartyChallanNo")
				PartyChallanDate = SetDateFormat(rsGRN("PartyChallanDate"))
				SecurityEntryNo = rsGRN("SecurityEntryNo")
				DeliveryDate = SetDateFormat(rsGRN("DeliveryDate"))
				LLRRNo = rsGRN("LLRRNo")
				VehicleNo = rsGRN("VehicleNo")
				Remarks = rsGRN("Remarks")				
				sSupName = rsGRN("SupplierName")
				sql= "Select SupplierAddress from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
				Call RunSql(sql,rsSup)
				if Not rsSup.Eof then
					sSupAddr = rsSup("SupplierAddress")
				end if
				rsSup.Close
		  %>
        <tr >
          <td bgcolor="#FFFFFF">
            <div align="center"><b>GRN No </b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Purchase Order No</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Purchase Request No</b></div>
          </td>
        </tr>
        <tr >
          <td bgcolor="#FFFFFF">
            <div align="center"><%=GetGRNNo(GRNNum)%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=GetPurchaseOrderNo(PurOrderNum)%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=GetPurchaseRequisitionNo(ReqNum)%></div>
          </td>
        </tr>
        <tr >
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Supplier Info </b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Party Challan No</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Party Challan Date</b></div>
          </td>
        </tr>
        <tr >
          <td vAlign="top" bgcolor="#FFFFFF">
            <div align="center"></div>
            <div align="center" >
              <% = sSupName %>
            </div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=PartyChallanNo%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=PartyChallanDate%></div>
          </td>
        </tr>
        <tr>
          <td rowspan="4" vAlign="top" style="word-break: break-all; width:350px;" align="center" bgcolor="#FFFFFF">
            <% =sSupAddr%>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Security Gate Entry No </b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Delivery Date</b></div>
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=SecurityEntryNo%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=DeliveryDate%></div>
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>LL/ RR No</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Vehicle No</b></div>
          </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=LLRRNo%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=VehicleNo%></div>
          </td>
        </tr>
        <%
				end if
				end if
				%>
        <tr>
          <td colspan="3" vAlign="top">
            <table width="100%" border="0" cellspacing="1" cellpadding="1">
              <tr >
                <td bgcolor="#FFFFFF">
                  <div align="center"><b>Sl.No</b></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><b>Item Description</b></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><b>Quantity Received</b></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><b>Quantity Accepted</b></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><b>Quantity Rejected</b></div>
                </td>
              </tr>
              <%
					sql = "Select ItemDescription,QtyReceived,QtyAccepted,QtyRejected from tbl_Psystem_GRN where GRNNo = "& GRnNo &" "
					Call runSql(sql,rsInfo)

					i = 1
					While Not rsInfo.Eof

					%>
              <tr  >
                <td bgcolor="#FFFFFF">
                  <div align="center"><%=i%></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><%=rsInfo("ItemDescription")%></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><%=rsInfo("QtyReceived")%></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><%=rsInfo("QtyAccepted")%></div>
                </td>
                <td bgcolor="#FFFFFF">
                  <div align="center"><%=rsInfo("QtyRejected")%></div>
                </td>
              </tr>
              <%
					i = i + 1
					rsInfo.movenext
					Wend
					rsInfo.Close

					%>
            </table>
          </td>
        </tr>
        <tr>
          <td colspan="3" bgcolor="#FFFFFF"><b>Remarks : </b> <%=Remarks%>
		  <p>&nbsp;</p><p>&nbsp;</p>
		  </td>
        </tr>
        <tr>
          <td colspan="3" bgcolor="#FFFFFF">Name &amp; Signature : </td>
        </tr>
        <tr>
          <td colspan="3" bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">Compiled By : </td>
          <td bgcolor="#FFFFFF">Verified By : </td>
          <td bgcolor="#FFFFFF">Authorised By :</td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF">&nbsp;</td>
          <td bgcolor="#FFFFFF">&nbsp;</td>
          <td bgcolor="#FFFFFF">&nbsp;</td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td align="Left" colspan="3" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" colspan="3" bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr>
    <td align="right">
      <INPUT  type="button" value="Print" class=formbutton  style="border: 1 solid" name=button2 onclick="window.print();">
    </td>
    <td align="Right">&nbsp;</td>
    <td align="Left">
      <input type="button" value="Close" class=formbutton  style="border: 1 solid" name=Close onclick="self.close();">
    </td>
  </tr>
  <tr>
    <td align="right">&nbsp;</td>
    <td align="Right">&nbsp;</td>
    <td align="Left">&nbsp;</td>
  </tr>
</table>



<!--#include file="../includes/connection_close.asp"-->