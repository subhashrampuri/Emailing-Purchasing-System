
<!--#include file="../includes/MailDesign.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/Connection_open.asp"-->
<!--#include file="../includes/style.asp"-->
<script type="text/javascript" src="toword.js">
</script>

<%
Function SetDateFormat(strDate)
'Response.Write(strDate)
dim arrDate
arrDate=split(strDate,"/")
SetDateFormat=arrDate(1)& "-" & MonthName(arrDate(0),true) & "-" & arrDate(2)
end function

	'Developed By: Subhash Rampuri
	Dim iReqID,sReqNo,sEName,sEId
	iReqID =  Request.QueryString("ReqID")
	PurOrdNo = Request.QueryString("PurOrdNo")
	sSupName = Request.QueryString("sSupplierName")
	
	sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrdNo &" "
	call RunSql(sql,rsPONum)
	if rsPONum.Eof = false then
		PurOrderNum = rsPONum("PurOrderNum")
	end if
	rsPONum.Close


	sql= "Select SupplierAddress,ContactPerson,TelephoneNo,MobileNo,EmailId,TINNo,ServiceTaxNo from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
	Call RunSql(sql,rsSup)

	if Not rsSup.Eof then
		SupplierAddr = rsSup("SupplierAddress")

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

%>
<title><%=gsSiteName%></title><table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td colspan="3" align="right"><img src="<%=gsSiteRoot%>gif/logo/Velankani_logo.gif" border = "0" ></td>
  </tr>
  <tr>
    <td align="center" colspan="3"><font size="5"><b>Purchase Order</b></font></td>
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
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>

  <tr>
    <td align="Left" colspan="3">
      <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
        <tr>
          <td  width="25%" bgcolor="#FFFFFF" valign="top">
            <div align="left">
              <p style="margin-left:5"><b>Supplier Name :</b></p>
            </div>
          </td>
          <td  width="25%" bgcolor="#FFFFFF" valign="top">
            <div align="left">
              <p style="margin-left:5"><%=sSupName%></p>
            </div>
          </td>
          <td  width="25%" bgcolor="#FFFFFF" valign="top">
            <div align="left">
              <p style="margin-left:5"><b>PurchaseOrder No :</b></p>
            </div>
          </td>
          <td  width="25%"  bgcolor="#FFFFFF"valign="top">
            <p style="margin-left:5"><%=GetPurchaseOrderNo(PurOrderNum)%></p>
          </td>
        </tr>
        <tr>
          <td  width="25%"  bgcolor="#FFFFFF"valign="top">
            <div align="left">
              <p style="margin-left:5"><b>Supplier Address :</b></p>
            </div>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%">
            <div align="left">
              <p style="margin-left:5">
                <% =SupplierAddr %>
              </p>
            </div>
          </td>
          <td  valign="top"  bgcolor="#FFFFFF" width="25%">
            <div align="left">
              <p style="margin-left:5"><b>Date:</b></p>
            </div>
          </td>
          <%
		  sql = "Select PurOrderDate from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrdNo &" "
		  call Runsql(sql,rsDate)

		  if rsDate.Eof = false then
		  	PurOrdDate = rsDate("PurOrderDate")
		  end if
		  rsDate.Close
		  %>
          <td  valign="top" bgcolor="#FFFFFF" width="25%">
            <p style="margin-left:5"><%=SetDateFormat(PurOrdDate)%></p>
          </td>
        </tr>
        <tr>
          <td  width="25%"  bgcolor="#FFFFFF"valign="top">
            <div align="left">
              <p style="margin-left:5"><b>Phone No :</b></p>
            </div>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%">
            <p style="margin-left:5"><%=TelephoneNo%></p>
          </td>
          <td  valign="top"  bgcolor="#FFFFFF" width="25%">
            <div align="left">
              <p style="margin-left:5"><b>Contact Person Name : </b></p>
            </div>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%">
            <p style="margin-left:5"><%=ContactPerson %></p>
          </td>
        </tr>
        <tr>
          <td  width="25%"  bgcolor="#FFFFFF"valign="top">
            <div align="left">
              <p style="margin-left:5"><b>Email ID :</b></p>
            </div>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%">
            <p style="margin-left:5"><%=EmailId %></p>
          </td>
          <td  valign="top"  bgcolor="#FFFFFF" width="25%">
            <div align="left">
              <p style="margin-left:5"><b>Contact No : </b></p>
            </div>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%">
            <p style="margin-left:5"><%=MobileNo %></p>
          </td>
        </tr>
        <tr>
          <td  width="25%"  bgcolor="#FFFFFF" valign="top">
            <p style="margin-left:5"><b>TIN No :</b></p>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%"><p style="margin-left:5"><%=TINNo %></p></td>
          <td  valign="top"  bgcolor="#FFFFFF" width="25%">
            <p style="margin-left:5"><b>Service Tax No :</b></p>
          </td>
          <td  valign="top" bgcolor="#FFFFFF" width="25%"><p style="margin-left:5"><%=ServiceTaxNo %></p></td>
        </tr>
      </table>
		<% rsSup.Close %>
    </td>
  </tr>
  <tr>
    <td align="center" colspan="3">
	  <table width="90%" cellpadding="2" cellspacing="2" align="center">
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><b>Dear Sir / Madam,</b></td>
        </tr>
        <tr>
          <td>With reference to the above, we are pleased to release our Purchase
            Order for the following items under the terms and conditions as mentioned
            below.</td>
        </tr>
      </table>
	</td>
  </tr>
  <tr>
    <td align="center" colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" colspan="3">
	  <table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor='#666666'>
        <tr >
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Sl.No</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>ItemDescription</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Tax Percent</b></div>
          </td>
          <td bgcolor="#FFFFFF"> <div align="center"><b>Tax</b></div></td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Quantity</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Unit Price</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><b>Amount</b></div>
          </td>
        </tr>
        <%
			'	if iReqID <> "" and sSupName <> "" then
				 sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and SupplierName = '" & sSupName & "' and RequisitionId = "& iReqID &" and PurOrderNo = "& PurOrdNo &""
				 call RunSql(sql,rsInfo)
				 i = 1
				 while not rsInfo.EOF
				 if rsInfo("Currency") = -1 then
					Curr = "Rs."
				 else
					Curr = "$"
				 end if
 				 if rsInfo("isTaxIncludedOrExcluded") = -1 then
					sTax = "Exclusive"
				 else
					sTax = "Inclusive"
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
        <tr >
          <td bgcolor="#FFFFFF">
            <div align="center"><%=i%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=rsInfo("ItemDescription")%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=rsInfo("TaxPercent") & " " & "%" %></div>
          </td>
          <td bgcolor="#FFFFFF">
		  <div align="center"><%=sTax%></div></td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=rsInfo("Quantity")%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=Curr & " " & rsInfo("UnitPrice")%></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=FormatNumber(Total,2)%></div>
          </td>
        </tr>
        <%
			  i = i + 1
			  GTotal = cDbl(GTotal) + cDbl(FormatNumber(Total,2))
			  rsInfo.movenext
			  Wend
			  rsInfo.Close
		'	end if
		  %>
        <tr>
          <td bgcolor="#FFFFFF"colspan="5">&nbsp;</td>
          <td bgcolor="#FFFFFF">
            <div align="right"><b>Grand Total :</b></div>
          </td>
          <td bgcolor="#FFFFFF">
            <div align="center"><%=Curr & " " & FormatNumber(GTotal) %></div>
          </td>
        </tr>
      </table>

	</td>
  </tr>
  <tr>
    <td align="center" colspan="3">
	  <table width="90%" cellpadding="2" cellspacing="2" align="center">
        <tr>

          <td>
		  <% X =  FormatNumber(GTotal,2)
			 Z = Replace(X,",","")
		   %>
		  <b>Amount in Words :</b>&nbsp;&nbsp;
		<script  type="text/javascript">
		//	alert(toWords(<%=Z %>));
			<%	if Curr = "Rs." then %>
			document.write (toWords(<%=Z %>)+ 'Rupess Only');
			<%	else	%>
			document.write (toWords(<%=Z %>)+ 'Dollars Only');
			<%	end if	%>
		</script>

		  </td>
        </tr>
      </table>
	</td>
  </tr>
  <tr>
  <%
  sql="Select RequiredDate, PaymentTerms from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrdNo &" "
  call RunSql(sql,rsPurOrd)
	if rsPurOrd.Eof = false then
		DeliveryDate = rsPurOrd("RequiredDate")
		PayTerms = rsPurOrd("PaymentTerms")
	end if
	rsPurOrd.Close
  %>

    <td align="center" colspan="3">
      <table width="90%" border="0" cellspacing="2" cellpadding="2">
        <tr>
          <td colspan="3"><b>Terms and Conditions </b></td>
        </tr>
        <tr>
          <td width="55%" colspan="2"><b>Delivery to :</b></td>
          <td width="35%">
            <p style="margin-left:50"><b>Velankani Software Pvt. Ltd., </b></p>
          </td>
        </tr>
        <tr>
          <td rowspan="5" colspan="2">&nbsp;</td>
          <td>
            <p style="margin-left:50">43, Electronic City Phase II</p>
          </td>
        </tr>
        <tr>
          <td>
            <p style="margin-left:50">Hosur Road Bangalore - 560 100</p>
          </td>
        </tr>
        <tr>
          <td>
            <p style="margin-left:50">Ph: 080 4037 5300</p>
          </td>
        </tr>
        <tr>
          <td>
            <p style="margin-left:50">Fax: 080 5514 5303</p>
          </td>
        </tr>
        <tr>
          <td>
            <p style="margin-left:50">Email: vspl_purchase@velankani.com</p>
          </td>
        </tr>
        <tr>
          <td colspan="2"><b>Date of Delivery : </b>
          </td>
          <td><p style="margin-left:50"><%=SetDateFormat(DeliveryDate)%> </p></td>
        </tr>
        <tr>
          <td colspan="3"><b>Taxes : </b></td>
        </tr>
        <tr>
          <td colspan="3"><b>Payment : </b> <%=PayTerms%></td>
        </tr>
        <tr>
          <td colspan="3"><b>Note : </b></td>
        </tr>
        <tr>
          <td colspan="3"><b>Bills : </b>Submit in Triplicate</td>
        </tr>
        <tr>
          <td colspan="3">&nbsp;</td>
        </tr>
        <tr>
          <td>
            <div align="center"><b>Verified By</b></div>
          </td>
          <td>
            <div align="center"><b>Approved By</b></div>
          </td>
          <td>
            <p style="margin-left:45"><b>For VELANKANI SOFTWARE PVT. LTD.,</b></p>
          </td>
        </tr>
        <tr>
          <td>
            <div align="center"></div>
          </td>
          <td>
            <div align="center"></div>
          </td>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td colspan="2">&nbsp;</td>
          <td>
            <p style="margin-left:45"><b>Authorized Signatory</b></p>
          </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td align="center" colspan="3">

	</td>
  </tr>
  <tr>
    <td  colspan="3">
      <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" >
        <tr >
          <td ><hr><b>Remarks : </b>
		  <p>&nbsp;</p><p>&nbsp;</p>
		  </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td  colspan="3">&nbsp;</td>
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
</table>
  <% X =  FormatNumber(GTotal,2)
		 Z = Replace(X,",","")
	   %>




<!--#include file="../includes/connection_close.asp"-->