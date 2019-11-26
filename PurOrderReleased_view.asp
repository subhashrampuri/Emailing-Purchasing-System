<%@ LANGUAGE="VBSCRIPT" %>
<%
'iMorfus Intranet Systems - Version 3.0.5 ' - Copyright 2002 - 04 (c) i-Vista Digital Solutions Limited.
'All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions.
'See the file iMorfuslicense.txt for more information.

'All Copyright notices must remain in place at all times.
'Developed By: Subhash Rampuri
'-----------------------------------------------------------------------------------------------

%>
<!--#include file="../includes/MailDesign.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
      <%
		PurOrderNo = Request.Form("hdPurOrdNo")

		sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrderNo &" "
		call RunSql(sql,rsPONum)
		if rsPONum.Eof = false then
			PurOrderNum = rsPONum("PurOrderNum")
		end if
		rsPONum.Close

	%>
	<script language="javascript">
	function CancelPO()
	{
		//alert(document.strFormm.hdPurOrdNo.value);
		document.strFormm.method="post";
		document.strFormm.action="CancelPO.asp";
		document.strFormm.submit();
	}
	</script>
	<form name="strFormm">


      <Table cellspacing="2" celpadding="2" align="center" width="100%">
        <tr class="blue">
          <td width="95%" height="25">
            <div align="center"><font color="#ffffff"><b>Released Purchase Orders</b></font></div>
          </td>
          <td width="5%" height="25">
            <div align="center"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></div>
          </td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td width="100%" colspan="2">
            <%
		  sql ="Select distinct SupplierName,PurOrderDate from tbl_Psystem_Quotations where PurOrderNo = "& PurOrderNo &" "
		  Call RunSql(sql,rsPSup)
		  sSupName = rsPSup("SupplierName")
     	  PurOrderDate = rsPSup("PurOrderDate")
			sql= "Select SupplierAddress,ContactPerson,TelephoneNo,MobileNo,EmailId,TINNo,ServiceTaxNo from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
			Call RunSql(sql,rsSup)

			if Not rsSup.Eof then
				sSupAddr = rsSup("SupplierAddress")
				if rsSup("TelephoneNo") <> "" then
					TelephoneNo  = rsSup("TelephoneNo")
				else
					TelephoneNo = ""
				end if

				if rsSup("ContactPerson") <> "" then
					ContactPerson  = rsSup("ContactPerson")
				else
					ContactPerson = ""
				end if

				if rsSup("EmailId") <> "" then
					EmailId  = rsSup("EmailId")
				else
					EmailId = ""
				end if

				if rsSup("MobileNo") <> "" then
					MobileNo  = rsSup("MobileNo")
				else
					MobileNo = ""
				end if

				if rsSup("ServiceTaxNo") <> "" then
					ServiceTaxNo  = rsSup("ServiceTaxNo")
				else
					ServiceTaxNo = ""
				end if

				if rsSup("TINNo") <> "" then
					TINNo  = rsSup("TINNo")
				else
					TINNo = ""
				end if

			end if
			rsSup.Close

		  %>
            <table width='100%' border='0' cellspacing='2' cellpadding='2'>
              <tr >
                <td colspan='6' align="right" ><a href="javascript:CallReqID();"><img src="images/printer.gif" width="20" height="21" border="0"></a></td>
              </tr>
              <tr >
                <td colspan='6' align="right" >
                  <table width="100%" border="0" cellspacing="2" cellpadding="2" align="center">
                    <tr>
                      <td class="blue" width="25%" valign="top">
                        <div align="right"><font color='#ffffff' ><b>Supplier
                          Name :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" width="25%" valign="top">
                        <div align="left"><%=sSupName%></div>
                      </td>
                      <td class="blue" width="25%" valign="top">
                        <div align="right"><font color='#ffffff' ><b>PurchaseOrder
                          No :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" width="25%" valign="top"><%=GetPurchaseOrderNo(PurOrderNum)%></td>
                    </tr>
                    <tr>
                      <td class="blue" width="25%" valign="top">
                        <div align="right"><font color='#ffffff' ><b>Supplier
                          Address :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%">
                        <div align="left">
                          <% =sSupAddr %>
                        </div>
                      </td>
                      <td class="blue" valign="top" width="25%">
                        <div align="right"><font color='#ffffff' ><b>Date:</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=SetDateFormat(PurOrderDate)%></td>
                    </tr>
                    <tr>
                      <td class="blue" width="25%" valign="top">
                        <div align="right"><font color='#ffffff' ><b>Phone No :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=TelephoneNo%></td>
                      <td class="blue" valign="top" width="25%">
                        <div align="right"><font color='#ffffff' ><b>Contact Person Name : </b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=ContactPerson %></td>
                    </tr>
                    <tr>
                      <td class="blue" width="25%" valign="top">
                        <div align="right"><font color='#ffffff' ><b>Email ID :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=EmailId %></td>
                      <td class="blue" valign="top" width="25%">
                        <div align="right"><font color='#ffffff' ><b>Contact No : </b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=MobileNo%></td>
                    </tr>
                    <tr>
                      <td class="blue" width="25%" valign="top">
                        <div align="right"><font color='#ffffff' ><b>TIN No :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=TINNo%></td>
                      <td class="blue" valign="top" width="25%">
                        <div align="right"><font color='#ffffff' ><b>Service Tax No :</b></font></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" valign="top" width="25%"><%=ServiceTaxNo%></td>
                    </tr>
                  </table>
                </td>
              </tr>
              <tr>
                <td colspan="6">&nbsp;</td>
              </tr>
              <tr class="blue">
                <td>
                  <div align='center'><font color='#ffffff' ><b>Sl.No</b></font></div>
                </td>
                <td>
                  <div align='center'><font color='#ffffff' ><b>ItemDescription</b></font></div>
                </td>
                <td>
                  <div align='center'><font  color='#ffffff'><b>Tax Percent</b></font></div>
                </td>
                <td>
                  <div align='center'><font  color='#ffffff'><b>Quantity</b></font></div>
                </td>
                <td>
                  <div align='center'><font  color='#ffffff'><b>Unit Price</b></font></div>
                </td>
                <td>
                  <div align='center'><font  color='#ffffff'><b>Amount</b></font></div>
                </td>
              </tr>
              <%
				if PurOrderNo <> "" and sSupName <> "" then
				 sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and SupplierName = '" & sSupName & "' and PurOrderNo = "& PurOrderNo &" "
				 call RunSql(sql,rsInfo)
				 i = 1
				 while not rsInfo.EOF
				 if rsInfo("Currency") = -1 then
					Curr = "Rs."
				 else
					Curr = "$"
				 end if
				 ReqId = rsInfo("RequisitionId")
				 Qty = rsInfo("Quantity")
				 Price = rsInfo("UnitPrice")
				 TaxPercent = rsInfo("TaxPercent")
				 Tax = (cDbl(TaxPercent) / 100)
				 Amount = (cInt(Qty) * cDbl(Price))
				 if rsInfo("isTaxIncludedOrExcluded") = -1 then
					Total = ((Amount) + ((Amount)* cDbl(Tax)))
				 else
					Total = Amount
				 end if
			%>
              <tr bgcolor="<%=gsBGColorLight%>">
                <td>
                  <div align='center'><%=i%></div>
                </td>
                <td style="word-break: break-all; width:300px;">
                  <div align='center'><%=rsInfo("ItemDescription")%></div>
                </td>
                <td>
                  <div align='center'><%=rsInfo("TaxPercent") & " " & "%" %></div>
                </td>
                <td>
                  <div align='center'><%=rsInfo("Quantity")%></div>
                </td>
                <td>
                  <div align='center'><%=Curr & " " & rsInfo("UnitPrice")%></div>
                </td>
                <td>
                  <div align='center'><%=FormatNumber(Total,2)%></div>
                </td>
              </tr>
              <%
				  i = i + 1
				  GTotal = cDbl(GTotal) + cDbl(FormatNumber(Total,2))
				  rsInfo.movenext
				  Wend
				  rsInfo.Close
				end if
			  %>
              <tr>
                <td colspan='4' bgcolor="<%=gsBGColorLight%>">&nbsp;</td>
                <td bgcolor="<%=gsBGColorLight%>">
                  <div align='right'><b>Grand Total: </b></div>
                </td>
                <td bgcolor="<%=gsBGColorLight%>">
                  <div align='center'><%=Curr & " " & FormatNumber(GTotal,2) %></div>
                </td>
              </tr>
              <tr>
                <td colspan='6'>&nbsp;</td>
              </tr>
              <tr>
                <td colspan='6' >&nbsp;</td>
              </tr>
              <tr>
                <td colspan='6' align="center">
				<%
				sql = "Select isGRNEntered from tbl_Psystem_Quotations where purOrderNo = "& PurOrderNo &" "
				call RunSql(sql,rsGRN)

				if rsGRN("isGRNEntered") = 0 then %>

				<input class=formbutton type=button value="Cancel PO" name="cancelpo" style="border: 1 solid" onclick="CancelPO();" >

				<% end if%>
				<input type="hidden" name="hdPurOrdNo" value="<%=PurOrderNo%>">
				</td>
              </tr>
            </table>
          </td>
        </tr>
        <tr>
          <td width="95%" colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td width="95%" colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td width="95%" colspan="2">&nbsp;</td>
        </tr>
        <tr bgcolor=#ffffff>
          <td colspan=5 align=center><A href="../../../main.asp"><%=dictLanguage("Return_Business_Console")%></a></td>
        </tr>
        <tr bgcolor=#ffffff>
          <td width="95%" colspan="2">&nbsp;</td>
        </tr>
      </table>
</table>

</form>
	<script language="JavaScript" type="text/javascript">
	function CallReqID()
	{
		myPopup = window.open("PurchaseOrder_Print.asp?ReqID=<%=ReqId%>&PurOrdNo=<%=PurOrderNo%>&sSupplierName=<%=sSupName%>",'popup','toolbar=no,width=800,height=600,scrollbars=yes,left=100,top=150,resizable=yes');
//		if (!myPopup.opener)
//	    myPopup.opener = self;
	}
	</script>

<br>

<!--#include file="../includes/connection_close.asp"-->
