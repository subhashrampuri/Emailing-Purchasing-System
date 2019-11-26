<%@ LANGUAGE="VBSCRIPT" %>
<%
'iMorfus Intranet Systems - Version 3.0.5 ' - Copyright 2002 - 04 (c) i-Vista Digital Solutions Limited.
'All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions.
'See the file iMorfuslicense.txt for more information.

'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------
'Developed By: Subhash Rampuri
'-----------------------------------------------------------------------------------------------

%>
<!--#include file="../includes/MailDesign.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->

    <%
 	  	ReqId = Request.Form("hdReqId")
		PurOrderNo = Request.Form("hdPurOrdNo")
		Supplier = Request.Form("hdSupplier")
		SupplierAddr = Request.Form("hdSupAddr")
		
		sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrderNo &" "
		call RunSql(sql,rsPONum)
		if rsPONum.Eof = false then
			PurOrderNum = rsPONum("PurOrderNum")
		end if
		rsPONum.Close
		

	%>
	<script language="JavaScript" type="text/javascript">
	function CallReqID()
	{
		myPopup = window.open('PurchaseOrder_Print.asp?ReqID=<%=ReqId %>&PurOrdNo=<%=PurOrderNo%>&sSupplierName=<%=Supplier%>',42,'toolbar=no,width=800,height=600,scrollbars=yes,left=100,top=150,resizable=yes');
		if (!myPopup.opener)
	    myPopup.opener = self;
	}
</script>

	<form name="strFormm">
      <table width="100%" align="center" valign="top" cellSpacing=2 cellPadding=2 width="100%" border=0>
  <Table cellspacing="2" celpadding="2" align="center" width="100%">
  <tr bgColor=#6699cc width="100%">
          <td bgcolor="#148ED3" class=homeheader align=middle >Purchase
            Order Released</td>
  </tr>
  <tr width="100%">
          <td align="right" ><a href="PurchaseTeamInbox.asp" style="text-decoration:none">Inbox</a></td>
  </tr>
  <tr valign="top">
          <td width="100%">
            <table width='100%' border='0' cellspacing='2' cellpadding='2'>
              <tr >
                <td colspan='6' align="right" ><a href="javascript:CallReqID();"><img src="images/printer.gif" width="20" height="21" border="0"></a>
                </td>
              </tr>
              <tr >
                <td colspan='6'>
                  <table width="100%" border="0" cellspacing="2" cellpadding="2">
                    <tr>
                      <td class="blue" width="25%" vAlign="top">
                        <div align="right"><b>Supplier Name :</b></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" width="25%" vAlign="top">
                        <div align="left"><%=Supplier%></div>
                      </td>
                      <td class="blue" width="25%" vAlign="top">
                        <div align="right"><b>PurchaseOrder No :</b></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" width="25%" vAlign="top"><%=GetPurchaseOrderNo(PurOrderNum)%></td>
                    </tr>
                    <tr>
                      <td class="blue" width="25%" vAlign="top">
                        <div align="right"><b>Supplier Address :</b></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" vAlign="top" width="25%">
                        <div align="left"><% =SupplierAddr %></div>
                      </td>
                      <td class="blue" vAlign="top" width="25%">
                        <div align="right"><b>Date:</b></div>
                      </td>
                      <td bgcolor="<%=gsBGColorLight%>" vAlign="top" width="25%"><%=SetDateFormat(Date())%></td>
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
                  <div align='center'><font color='#ffffff' ><b>Item Description</b></font></div>
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
				if ReqId <> "" and Supplier <> "" then
				 sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and SupplierName = '" & Supplier & "' and RequisitionId = "& ReqId &" and PurOrderNo = "& PurOrderNo &""
				 call RunSql(sql,rsInfo)
				 i = 1
				 while not rsInfo.EOF
				 if rsInfo("Currency") = -1 then
					Curr = "Rs."
				 else
					Curr = "$"
				 end if
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
                <td>
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
            </table>
          </td>
  </tr>
  <tr>
          <td width="95%">&nbsp;</td>
  </tr>
  <tr bgcolor=#ffffff>
    <td colspan=4 align=center><a href="PurchaseOrder.asp" style="text-decoration:none"><b>Back</b></a> <br><br> <A href="../../../main.asp"><%=dictLanguage("Return_Business_Console")%></a></td>
  </tr>
  <tr bgcolor=#ffffff>
          <td width="95%">&nbsp;</td>
  </tr>
</table>
</table>
</form>

<br>

<!--#include file="../includes/connection_close.asp"-->
