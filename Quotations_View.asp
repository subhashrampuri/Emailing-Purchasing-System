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
 <Script language=JavaScript src="../includes/javascript/validate.js" type=text/javascript></SCRIPT>
 <script language="JavaScript" type="text/javascript">
  function submit_approve()
	{
	var x=document.strFormm.elements.length;
	for(i=0;i<x;i++)
	if(document.strFormm.elements[i].type=="radio")
	 if(document.strFormm.elements[i].checked==true)
	{
		if(document.FinalForm.ItemList.value=="")
			document.FinalForm.ItemList.value=document.strFormm.elements[i].value;
		else
			document.FinalForm.ItemList.value=document.FinalForm.ItemList.value + "|" + document.strFormm.elements[i].value;
	}

	document.FinalForm.method="post";
	document.FinalForm.action="Quotation_Approvers.asp";
	document.FinalForm.submit();

  }

  function submit_reject()
	{
	var x=document.strFormm.elements.length;

	for(i=0;i<x;i++)
	if(document.strFormm.elements[i].type=="radio")
		if(document.strFormm.elements[i].checked==true)
		{
			if(document.FinalForm.ItemList.value=="")
				document.FinalForm.ItemList.value=document.strFormm.elements[i].value;
			else
				document.FinalForm.ItemList.value=document.FinalForm.ItemList.value + "|" + document.strFormm.elements[i].value;
		}
	//alert(document.FinalForm.hdReqNo.value);
	document.FinalForm.method="post";
	document.FinalForm.action="Quotation_Reject.asp";
	document.FinalForm.submit();

  }

function submit_onhold()
	{
	var x=document.strFormm.elements.length;

	for(i=0;i<x;i++)
	if(document.strFormm.elements[i].type=="radio")
		if(document.strFormm.elements[i].checked==true)
		{
			if(document.FinalForm.ItemList.value=="")
				document.FinalForm.ItemList.value=document.strFormm.elements[i].value;
			else
				document.FinalForm.ItemList.value=document.FinalForm.ItemList.value + "|" + document.strFormm.elements[i].value;
		}
	//alert(document.FinalForm.hdReqNo.value);
	document.FinalForm.method="post";
	document.FinalForm.action="Quotation_OnHold.asp";
	document.FinalForm.submit();

  }

</script>
	<%
		PurRequisitionNo = Request.Form("hdReqId")
		
		sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& PurRequisitionNo &" "
		call RunSql(sql,rsRec)
		if rsRec.Eof = false then
			ReqNum = rsRec("RequisitionNum")
		end if
		rsRec.Close
		
	%>
<form name="strFormm">
  <table width="100%" align="center" valign="top" cellSpacing=2 cellPadding=2 width="100%" border=0>
   <tr>
    <td>
      <table width="100%" cellpadding="1" cellspacing="2">
        <tr class="blue">
          <td colspan=9 align="center" width="95%" height="25"><font color=#ffffff><b>Approve
            / Reject Purchase Quotations</b></font></td>
          <td align="center" width="5%" height="25" ><a href="ApproverInbox.asp" style="text-decoration:none"><font color=#ffffff>Inbox</font></a></td>
        </tr>
        <tr>
          <td colspan=10>&nbsp; </td>
        </tr>
        <%
	  Dim i
	   Sql= "select distinct ItemDescription,ProjectId,RequisitionId from tbl_Psystem_Quotations where RequisitionId= "& PurRequisitionNo &" And isApproved = 0 and isPRCancelled = 0 "
	  	call RunSql(sql,rsItems)

		while Not rsItems.EOF

			iDesc=rsItems("ItemDescription")
			PrjId = rsItems("ProjectId")
			sql = sql_GetProjectName(PrjID)
			call RunSql(sql,rsPrj)
			if not rsPrj.EOF then
				sPrjName = rsPrj("ProjectName")
				rsPrj.Close
			end if

	  %>
        <tr >
          <td align="center" colspan="10">
            <table width="100%" border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="17%">
                  <input type="hidden" name="hdItemDesc" value="<%=iDesc%>_<%=PrjId%>">
                  <div align="right"><b>Requisition No : </b></div>
                </td>
                <td width="17%">&nbsp;<%=GetPurchaseRequisitionNo(ReqNum)%></td>
                <td width="17%">
                  <div align="right"><b>Item Description :</b></div>
                </td>
                <td width="17%" nowrap>&nbsp;<%=iDesc%></td>
                <td width="17%">
                  <div align="right"><b>Project :</b></div>
                </td>
                <td width="15%" nowrap>&nbsp;<%=sPrjName%></td>
              </tr>
            </table>
          </td>
        </tr>
        <tr class="blue">
          <td align="center"><font color=#ffffff><b>Sl. No.</b></font></td>
          <td align="center"><font color=#ffffff><b>Supplier Name</b></font></td>
          <td align="center"><font color=#ffffff><b>Unit Price</b></font></td>
          <td align="center"><font color=#ffffff><b>Tax</b></font></td>
          <td align="center"><font color=#ffffff><b>Tax Percent</b></font></td>
          <td align="center"><font color=#ffffff><b>Quantity</b></font></td>
          <td align="center"><font color=#ffffff><b>Warranty</b></font></td>
          <td align="center"><font color=#ffffff><b>Delivery Time</b></font></td>
          <td align="center"><font color=#ffffff><b>Payment Terms</b></font></td>
          <td align="center"><font color=#ffffff><b>Remarks</b></font></td>
        </tr>
        <%

		sql= "Select * from tbl_Psystem_Quotations where ItemDescription = '" & Replace(Server.HTMLEncode(iDesc),"'","''") & "' and ProjectId = "& PrjId &" and RequisitionId = "& PurRequisitionNo &" And isApproved = 0 and isPRCancelled = 0 "
		Call RunSql(sql,rsInfo)
		'if i <> 1 then
		'	i = i - 1
		'end if
		i = 1
		while NOT rsInfo.EOF
			if rsInfo("Currency") = -1 then
				sCurr = "Rupee(s)"
			else
				sCurr = "Doller(s)"
			end if
			if rsInfo("isTaxIncludedOrExcluded") = -1 then
				sTax = "Exclusive"
			else
				sTax = "Inclusive"
			end if
			iCode = rsInfo("ItemCode")

		%>
        <tr bgcolor="<%=gsBGColorLight%>">
          <td align="center">
            <input type="radio" name="<%=iDesc%>_<%=PrjId%>" value="<%=iCode%>" >
            <%=i%></td>
          <td align="left" style="word-break: break-all; width:200px;" vAlign="top"><%=rsInfo("SupplierName")%></td>
          <td align="center" vAlign="top"><%=rsInfo("UnitPrice") & " " & sCurr%></td>
          <td align="center" vAlign="top"><%=sTax%></td>
          <td align="center" vAlign="top"><%=rsInfo("TaxPercent") & " " & "%"%></td>
          <td align="center" vAlign="top"><%=rsInfo("Quantity")%></td>
          <td align="center" vAlign="top"><%=rsInfo("Warranty")%></td>
          <td align="center" vAlign="top"><%=rsInfo("DeliveryTime")%></td>
          <td align="center" vAlign="top"><%=rsInfo("PaymentTerms")%></td>
          <td align="center" vAlign="top"><%=rsInfo("Remarks")%></td>
        </tr>
        <%
			i = i + 1
			rsInfo.movenext
			Wend
			rsInfo.Close
		%>
        <tr>
          <td colspan=10>&nbsp; </td>
        </tr>
        <%
			'i = i + 1
			rsItems.movenext
			wend
			'iCount = i - 2
			rsItems.close

		%>
      </table>
    </td>
  </tr>

  <tr>
  <td>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor=#ffffff width=100%>
      <table width="20%" align="center" cellspacing=2 cellpadding=2 border=0>
        <tr height=25>
          <td bgcolor="<%=lclstr_bgColor%>" align=center>
            <input class=formbutton type=button value="Approve" name="Approve" style="border: 1 solid" onclick="submit_approve();">
          </td>
          <td width=60>&nbsp;</td>
          <td bgcolor="<%=lclstr_bgColor%>" align=center>
            <input class=formbutton type=button value="Reject PR" name="Reject" style="border: 1 solid" onclick="submit_reject();" >
          </td>
          <td bgcolor="<%=lclstr_bgColor%>" align=center>&nbsp;</td>
          <td bgcolor="<%=lclstr_bgColor%>" align=center>
            <input class=formbutton type=button value="On Hold PR" name="OnHold" style="border: 1 solid" onclick="submit_onhold();" >
          </td>
          <td> </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor=#ffffff>
    <td>&nbsp;</td>
  </tr>
  <tr bgcolor=#ffffff>
    <td  align=center><A href="../../../main.asp"><%=dictLanguage("Return_Business_Console")%></a></td>
  </tr>
  <tr bgcolor=#ffffff>
    <td>&nbsp;</td>
  </tr>
</table>
</form>
<form name="FinalForm">
<input type="hidden" name="ItemList" value="">
<input type="hidden" name="hdReqNo" value="<%=PurRequisitionNo%>">
</form>
<br>

<!--#include file="../includes/connection_close.asp"-->
