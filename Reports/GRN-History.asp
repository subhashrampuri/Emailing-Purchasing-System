<% Response.Buffer = True %>
<!--#include file="../../includes/main_page_header.asp"-->
<!--#include file="../../includes/main_page_open.asp"-->

<DIV ID="splashScreen" STYLE="position:absolute;z-index:5;top:30%;left:35%;">
	<TABLE BORDER=0 BORDERCOLOR="#148ED3" CELLPADDING=0 CELLSPACING=0 HEIGHT=150 WIDTH=150>
		<TR>
			<TD WIDTH="100%" HEIGHT="100%"  ALIGN="CENTER" VALIGN="MIDDLE">
				<img src="../images/loading.gif"><br><br>
			</TD>
		</TR>
	</TABLE>
</DIV>
<%Response.Flush%>
<%
	'Developed By: Subhash Rampuri
	
	FromDate = Request.Form("txtFromDate")
	ToDate = Request.Form("txtToDate")	
	Response.write FromDate & "-" & ToDate
	PurOrdNo = Request.Form("PurOrdNo")
	iGRNNo = Request.Form("GRNNo")
%>
<title><%=gsSiteName%></title>
<Script language=JavaScript src="../../includes/javascript/validate.js" type=text/javascript></SCRIPT>
<script language="javascript">
	function Validator(frm)
	{

		if(frm.PurOrdNo.value == "0")
		{
			alert("Please select purchase order");
			frm.PurOrdNo.focus();
			return false;
		}
		if(frm.GRNNo.value == "0")
		{
			alert("Please select GRN No");
			frm.GRNNo.focus();
			return false;
		}

		return true;
	}
	function export_excel()
	{
	//alert(document.Export_PO.hdFromDate.value);
	//alert(document.Export_PO.hdToDate.value);
	document.Export_GRN.method="Post";
	document.Export_GRN.action="Report_GRN-History.asp"
	document.Export_GRN.submit();

	}
	function ItemInfo()
	{
		
		document.strFormm.method="post";
		document.strFormm.action="GRN-History.asp";
		document.strFormm.submit();
	}
</script>
	<SCRIPT LANGUAGE="JavaScript">
		// This script is intended for use with a minimum of Netscape 4 or IE 4.
		if(document.getElementById) {
			var upLevel = true;
			}
		else if(document.layers) {
			var ns4 = true;
			}
		else if(document.all) {
			var ie4 = true;
			}

		function showObject(obj) {
			if (ns4) obj.visibility = "show";
			else if (ie4 || upLevel) obj.style.visibility = "visible";
			}
		function hideObject(obj) {
			if (ns4) {
				obj.visibility = "hide";
				}
			if (ie4 || upLevel) {
				obj.style.visibility = "hidden";
				}
			}
	</SCRIPT>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td align="center" colspan="3"><font size="4"><b>Goods Received Note</b></font></td>
  </tr>
  <tr>
    <td align="center" colspan="3">
      <hr>
    </td>
  </tr>
  <tr>
    <td  colspan="3"> </td>
  </tr>
  <tr>
    <td  colspan="3">&nbsp; </td>
  </tr>
  <tr>
    <td  colspan="3">
      <form name="strFormm" method="post" action="" onSubmit='return Validator(this)'>
        <table width="50%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#999999">
          <tr> 
            <td bgcolor="#FFFFFF" colspan="2" >&nbsp;</td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF"> 
              <div align="right"><b>Purchase Order No:</b></div>
            </td>
            <td bgcolor="#FFFFFF">&nbsp; 
              <Select class="formstylemed" name="PurOrdNo" onChange="ItemInfo(this.value)">
                <option Selected value="0">Select Purchase Released List</option>
                <%
					Sql =  " Select distinct PurOrderNo from tbl_Psystem_GRN Order by PurOrderNo " 
					Call RunSql(sql,rsList)

					While Not rsList.EOF
					PONo = rsList("PurOrderNo")

					sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PONo &" "
					call RunSql(sql,rsPONum)
					if rsPONum.Eof = false then
						PurOrderNum = rsPONum("PurOrderNum")
					end if
					rsPONum.Close

						if cInt(rsList("PurOrderNo")) = cInt(PurOrdNo) then
			 		%>
                <option Selected value="<%=rsList("PurOrderNo")%>"><%=GetPurchaseOrderNo(PurOrderNum)%></option>
                <% else %>
                <option value="<%=rsList("PurOrderNo")%>"><%=GetPurchaseOrderNo(PurOrderNum)%></option>
                <% end if %>
                <%
				 	rsList.movenext
					Wend
					rsList.close
				  %>
              </select>
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF"> 
              <div align="right"><b>GRN No:</b></div>
            </td>
            <td bgcolor="#FFFFFF">&nbsp;
                      <Select class="formstylemed" name="GRNNo" >
	                <option Selected value="0">Select GRN No</option>
                <%
				
				  if PurOrdNo <> "" then
					sql = " Select distinct GRNNo,GRNNum from tbl_Psystem_GRN where PurOrderNo = "& PurOrdNo &" "
					Call RunSql(sql,rsGRNPO)
					
					While Not rsGRNPO.EOF
					GRNNum = rsGRNPO("GRNNum")
					if cInt(rsGRNPO("GRNNo")) = cInt(iGRNNo) then
		 		%>
					<option Selected value="<%=rsGRNPO("GRNNo")%>"><%=GetGRNNo(GRNNum)%></option>
					<% else %>
					<option value="<%=rsGRNPO("GRNNo")%>"><%=GetGRNNo(GRNNum)%></option>
					<% end if %>
                <%
					rsGRNPO.movenext
					Wend
					rsGRNPO.close
				end if
				  %>
              </select>
			   
            </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td bgcolor="#FFFFFF"> &nbsp;&nbsp;&nbsp;&nbsp; 
              <input type="submit" name="Submit" value="Submit">
              &nbsp;&nbsp; </td>
          </tr>
          <tr> 
            <td bgcolor="#FFFFFF" colspan="2">&nbsp;</td>
          </tr>
        </table>
        <br>
      </form>
    </td>
  </tr>
  <tr>
    <td  colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td  colspan="3">
	<%
		'response.write PurOrdNo  & iGRNNo
		if  PurOrdNo <> 0 and iGRNNo <> 0 then

	%>

	<table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
        <tr>
          <td align="Center" width="90%" bgcolor="#FFFFFF">&nbsp;

          </td>
          <td width="10%" bgcolor="#FFFFFF">
		  <input type="button" name="Export" value="Export-Excel" onClick="javascript:export_excel()">
		  </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" colspan="2"> <br>
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
            <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#999999">
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
            <br>
          </td>
        </tr>
      </table>
      <% end if %>
	</td>
  </tr>
  <tr>
    <td  colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td  colspan="3">
	<form name="Export_GRN">
	<input type="hidden" name="hdPurOrdNo" value ="<%=PurOrdNo%>">
	<input type="hidden" name="hdGRNNo" value ="<%=iGRNNo%>">	
	</form>

	</td>
  </tr>
</table>
	<SCRIPT LANGUAGE="JavaScript">
	if(upLevel) {
		var splash = document.getElementById("splashScreen");
		}
	else if(ns4) {
		var splash = document.splashScreen;
		}
	else if(ie4) {
		var splash = document.all.splashScreen;
		}
	hideObject(splash);
	</SCRIPT>





<!--#include file="../../includes/connection_close.asp"-->