
<%
	Response.Buffer =True
	Server.ScriptTimeout = 10000
	Response.ContentType = "application/vnd.ms-excel"

	Response.Charset = "GB2312"
	Response.Codepage = "936"

	FromDate = Trim(Request.Form("hdFromDate"))
	ToDate = Trim(Request.Form("hdToDate"))

%>

<!--#include file="../../includes/main_page_header.asp"-->

	<%
	sql = " select distinct dbo.fn_Psystem_PurchaseOrderNo(a.PurOrderNum)as PurOrderNo,b.PurOrderNo as PurOrdNo,b.SupplierName, " & _
		" dbo.fn_Psystem_GetSupplierAddress(b.SupplierName) as SupplierAddress,dbo.fn_TSystem_GetVelankaniFormatDate(a.PurOrderDate)as PurOrderDate " & _
		" from tbl_Psystem_PurchaseOrder a, tbl_Psystem_Quotations b where a.PurOrderNo = b.PurOrderNo and a.PurOrderDate between  " & _ 
		" '" & FromDate & "' and '" & ToDate & "'  and isPOCancelled = 1 "
	Call RunSql(sql,objRs)
	
	Response.AddHeader "content-disposition","attachment;filename=Purchase_Order_Cancelled" & FromDate & "_" & ToDate & ".xls"
	
	if objRs.EOF = false then
		While NOT objRs.EOF
		PurOrdNo = objRs("PurOrdNo")
		SupplierName = objRs("SupplierName")
	Sql = "select SupplierAddress,TINNo,ServiceTaxNo from tbl_Psystem_Supplier where SupplierName = '" & SupplierName & "'"
	Call RunSql(sql,rsSup)
	
	if rsSup("TINNo") <> "" then
		TINNo = rsSup("TINNo")
	else
		TINNo =""
	end if

	if rsSup("ServiceTaxNo") <> "" then
		ServiceTaxNo = rsSup("ServiceTaxNo")
	else
		ServiceTaxNo =""
	end if
	rsSup.Close
	
	%>

            
<table width="90%" border="1" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
  <tr valign="top"> 
                <td  bgcolor="#FFFFFF" nowrap> 
                  <p style="margin-left:10"><b>Purchase Order No :</b></p>
                </td>
                <td  bgcolor="#FFFFFF"> 
                  <p style="margin-left:10"><%=objRs("PurOrderNo")%></p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap> 
                  <p style="margin-left:10"><b>Supplier Name :</b></p>
                </td>
                <td bgcolor="#FFFFFF"> 
                  <p style="margin-left:10"><%=SupplierName%> </p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap colspan="2">
				  <p style="margin-left:10"><b>Supplier TIN No :</b></p>
				</td>
                <td  bgcolor="#FFFFFF">
				<p style="margin-left:10"><%=TINNo%></p>
				</td>
              </tr>
              <tr valign="top"> 
                <td  bgcolor="#FFFFFF" nowrap> 
                  <p style="margin-left:10"><b>Purchase Order Date :</b></p>
                </td>
                <td  bgcolor="#FFFFFF"> 
                  <p style="margin-left:10"><%=objRs("PurOrderDate")%></p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap> 
                  <p style="margin-left:10"><b>Supplier Address : </b></p>
                </td>
                <td  bgcolor="#FFFFFF"> 
                  <p style="margin-left:10"><%=objRs("SupplierAddress")%> </p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap colspan="2">
				<p style="margin-left:10"><b>Supplier Tax No :</b></p></td>
                <td  bgcolor="#FFFFFF"><p style="margin-left:10"><%=ServiceTaxNo%></p></td>
              </tr>
            </table>
            
<table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor='#666666' border="1">
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
                <td bgcolor="#FFFFFF"> 
                  <div align="center"><b>Tax</b></div>
                </td>
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
			 sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and SupplierName = '" & SupplierName & "' and PurOrderNo = "& PurOrdNo &""
			 call RunSql(sql,rsInfo)
			 i = 1
			 GTotal = 0
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
                  <div align="center"><%=sTax%></div>
                </td>
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
            <br>
            <br>
            <%
		objRs.movenext
		Wend
		end if
	%>


<!--#include file="../../includes/connection_close.asp"-->