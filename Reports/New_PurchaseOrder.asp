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
	FromDate = Trim(Request.form("txtFromDate"))
	ToDate = Trim(Request.form("txtToDate"))

%>
<title><%=gsSiteName%></title>
<Script language=JavaScript src="../../includes/javascript/validate.js" type=text/javascript></SCRIPT>
<script language="javascript">
	function Validator(frm)
	{
		var str,s,i
    	formElements=["txtFromDate","txtToDate"];
     	for(i=0;i<1;i++)
    	{
	      if(frm.elements[formElements[i]].value.length !=0)
    	  {
        	 str=frm.elements[formElements[i]].value
	         s = str.replace(/^(\s)*/, '');
	         s = s.replace(/(\s)*$/, '');
	         frm.elements[formElements[i]].value=s
    	  }
	    }
		if(frm.txtFromDate.value == "")
		{
			alert("Please enter from date");
			frm.txtFromDate.focus();
			return false;
		}
		if(frm.txtToDate.value == "")
		{
			alert("Please enter to date");
			frm.txtToDate.focus();
			return false;
		}
		else if (isGreaterDate(ChangeToMMDDYYYY(document.strFormm.txtToDate.value),ChangeToMMDDYYYY(document.strFormm.txtFromDate.value)))
		{
			alert("To date cannot be less than from date");
			document.strFormm.txtToDate.focus();
			return false;
		}
		return true;
	}
	function export_excel()
	{
	//alert(document.Export_PO.hdFromDate.value);
	//alert(document.Export_PO.hdToDate.value);
	document.Export_PO.method="Post";
	document.Export_PO.action="Report_PurchaseOrder.asp"
	document.Export_PO.submit();

	}
		function ChangeToMMDDYYYY(strDate)
		{
			var datarr;
			var strmon;
			datarr=strDate.split("-")
			switch(datarr[1])
			{
				case 'Jan':
					strmon="01";
					break;
				case 'Feb':
					strmon="02";
					break;
				case 'Mar':
					strmon="03";
					break;
				case 'Apr':
					strmon="04";
					break;
				case 'May':
					strmon="05";
					break;
				case 'Jun':
					strmon="06";
					break;
				case 'Jul':
					strmon="07";
					break;
				case 'Aug':
					strmon="08";
					break;
				case 'Sep':
					strmon="09";
					break;
				case 'Oct':
					strmon="10";
					break;
				case 'Nov':
					strmon="11";
					break;
				case 'Dec':
					strmon="12";
					break;
			}
			if(datarr[0]=="1"||datarr[0]=="2"||datarr[0]=="3"||datarr[0]=="4"||datarr[0]=="5"||datarr[0]=="6"||datarr[0]=="7"||datarr[0]=="8"||datarr[0]=="9")
				datarr[0]="0"+datarr[0];

			return (strmon+'/'+datarr[0]+'/'+datarr[2]);
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
    <td colspan="3" align="right">&nbsp;</td>
  </tr>
  <tr>
    <td align="center" colspan="3"><font size="4"><b>Purchase Order</b></font></td>
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
              <div align="right"><b>From Date:</b></div>
            </td>
            <td bgcolor="#FFFFFF"> &nbsp;&nbsp;
              <input type="text" class="formstylemedium" name="txtFromDate" Readonly>
              &nbsp;&nbsp;<a name="CalanderLink" onClick="openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','txtFromDate',150,300)"><img border="0" src="../../gif/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" width="16" height="15"></a>
            </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">
              <div align="right"><b>To Date:</b></div>
            </td>
            <td bgcolor="#FFFFFF"> &nbsp;&nbsp;
              <input type="text" class="formstylemedium" name="txtToDate" Readonly>
              &nbsp;&nbsp;<a name="CalanderLink" onClick="openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','txtToDate',150,300)"><img border="0" src="../../gif/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" width="16" height="15"></a>
            </td>
          </tr>
          <tr>
            <td bgcolor="#FFFFFF">&nbsp;</td>
            <td bgcolor="#FFFFFF"> &nbsp;&nbsp;&nbsp;&nbsp;
              <input type="submit" name="Submit" value="Submit">
              &nbsp;&nbsp;
              <input type="reset" name="Reset" value="Reset">
              <input type="hidden" name="hdFromDate" value="<%=FromDate%>">
              <input type="hidden" name="hdToDate" value="<%=ToDate%>">
            </td>
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
	'Response.write FromDate & " " & ToDate
	if FromDate <> "" and ToDate <> "" then
	%>
      <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
        <tr>
          <td align="center" bgcolor="#FFFFFF" width="90%">
            <% Response.write "Purchase Order between <b>" &  FromDate & "</b> and <b>" & ToDate %>
          </td>
          <td align="center" bgcolor="#FFFFFF" width="10%">
		  <input type="button" name="Export" value="Export-Excel" onClick="javascript:export_excel()">
		  </td>
        </tr>
        <tr>
          <td bgcolor="#FFFFFF" colspan="2">
            <%

	sql = " select distinct dbo.fn_Psystem_PurchaseOrderNo(a.PurOrderNum)as PurOrderNo,b.PurOrderNo as PurOrdNo,b.SupplierName, " & _
		" dbo.fn_Psystem_GetSupplierAddress(b.SupplierName) as SupplierAddress,dbo.fn_TSystem_GetVelankaniFormatDate(a.PurOrderDate)as PurOrderDate " & _
		" from tbl_Psystem_PurchaseOrder a, tbl_Psystem_Quotations b where a.PurOrderNo = b.PurOrderNo and a.PurOrderDate between  " & _
		"  '" & FromDate & "' and '" & ToDate & "'  and b.isClosed = 0 Order by b.PurOrderNo "

		Call RunSql(sql,objRs)

	if objRs.EOF = false then
		While NOT objRs.EOF
		PurOrdNo = objRs("PurOrdNo")
		SupplierName = objRs("SupplierName")
	Sql = "select SupplierAddress,TINNo,ServiceTaxNo from tbl_Psystem_Supplier where SupplierName = '" & SupplierName & "'"
	Call RunSql(sql,rsSup)

	if rsSup("TINNo") <> "" then
		TINNo = rsSup("TINNo")
	else
		TINNo = ""
	end if

	if rsSup("ServiceTaxNo") <> "" then
		ServiceTaxNo = rsSup("ServiceTaxNo")
	else
		ServiceTaxNo =""
	end if
	rsSup.Close

	%>
            <br>
            <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
              <tr valign="top">
                <td  bgcolor="#FFFFFF" nowrap>
                  <p style="margin-left:10"><b>Purchase Order No :</b></p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap>
                  <p style="margin-left:10"><%=objRs("PurOrderNo")%></p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap>
                  <p style="margin-left:10"><b>Supplier Name :</b></p>
                </td>
                <td bgcolor="#FFFFFF">
                  <p style="margin-left:10"><%=SupplierName%> </p>
                </td>
                <td  bgcolor="#FFFFFF" nowrap>
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
                <td  bgcolor="#FFFFFF" nowrap>
                  <p style="margin-left:10"><b>Supplier Tax No :</b></p>
                </td>
                <td  bgcolor="#FFFFFF">
                  <p style="margin-left:10"><%=ServiceTaxNo%></p>
                </td>
              </tr>
            </table>
            <table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor='#666666' border="0">
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
		objRs.close
		else
		Response. Write "<tr><td align='center' bgcolor='#ffffff' Colspan='2'><b>No Records Found</b></td></tr>"
		end if
	%>
          </td>
        </tr>
      </table>
      <% end if%>
    </td>
  </tr>
  <tr>
    <td  colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td  colspan="3">
	<form name="Export_PO">
	<input type="hidden" name="hdFromDate" value="<%=FromDate%>">
	<input type="hidden" name="hdToDate" value="<%=ToDate%>">
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