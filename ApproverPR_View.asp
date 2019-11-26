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
	var p=0;
	var half=0;
	var cnt=0;
	//alert(x);
	for(i=0;i<x;i++)
	{
	if(document.strFormm.elements[i].type=="checkbox")
		{
		p=p+1;

		}
	}
	half=p/2;
	for(k=1;k<=half;k++)
	{
		if (document.getElementById("action_"+k).checked == false)
		  {
			cnt=cnt+1;
		  }
	} //for loop

	  if (half==cnt)
	  {
	   alert('Select an item to approve');
	   return false;
	  }

	document.strFormm.method="post";
	document.strFormm.action="submit_Approvers.asp";
	document.strFormm.submit();

  }

  function submit_reject()
	{
	var x=document.strFormm.elements.length;
	var p=0;
	var half=0;
	var cnt=0;
	//alert(x);
	for(i=0;i<x;i++)
	{
	if(document.strFormm.elements[i].type=="checkbox")
		{
		p=p+1;

		}
	}
	half=p/2;
	for(k=1;k<=half;k++)
	{
		if (document.getElementById("action_"+k).checked == false)
		  {
			cnt=cnt+1;
		  }
	} //for loop

	  if (half==cnt)
	  {
	   alert('Select an item to reject');
	   return false;
	  }

	document.strFormm.method="post";
	document.strFormm.action="submit_Reject.asp";
	document.strFormm.submit();
  }

</script>
<script language="Javascript">
	var isToPropagate=true;
  function Check(k)
	{
	//alert(k);
	if (document.getElementById("action_"+k).checked == true)
		{

		document.getElementById("Quantity_"+k).disabled = false;
		document.getElementById("Ownby_"+k).disabled = false;
		}
	else
		{
		document.getElementById("Quantity_"+k).disabled = true;
		document.getElementById("Ownby_"+k).disabled = true;
		}
	}
	function validateQuantity(ctrl)
	{

		var regexp = new RegExp (/^[0-9]\d*$/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		var  a= ctrl.name;
		var indexarr = a.split("_");
		var Num = indexarr[1];
		var AQty = document.getElementById("hdQty" + Num);
		var N  = ctrl.value;
		function significantnumber( N, significance )
		{
		  if(typeof significance==='undefined') { significance = 2; }
		  N = Math.round( N * Math.pow( 10, significance ) ) / Math.pow( 10, significance );
		  return N; // may need padding
		}
		//alert(N);
		if(isToPropagate==true)
		{
		 if(!regexp.test(ctrl.value))
			{
				alert("Please enter a valid Quantity Required.");
				isToPropagate=false;
				ctrl.value = AQty.value;
				ctrl.focus();
				return false;
			}
	//		else if (parseInt(ctrl.value) > parseInt(AQty.value))
			else if (N > parseInt(AQty.value))
			{
				alert("Quantity cannot be increased");
				isToPropagate=false;
				ctrl.value = AQty.value;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;
		}
		return true;
	}
</script>
<form name="strFormm">
  <table width="100%" align="center" valign="top" cellSpacing=2 cellPadding=2  border=0>

  <tr>
    <td>
      <table width="100%" align="center" cellSpacing=2 cellPadding=2 border=0>
        <tr class="blue">
          <td align="center" colspan="9" width="95%"><font color="#ffffff"><b>Approve / Reject Purchase
            Request</b></font></td>
          <td align="center" width="5%"><a href="ApproverInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></td>
        </tr>
        <tr >
          <td align="center" colspan="10" >&nbsp;</td>
        </tr>
        <tr >
          <td align="center" colspan="10" >&nbsp;</td>
        </tr>

        <tr class="blue">
          <td align="center" ><font color=#ffffff><b>Sl. No.</b></font></td>
          <td align="center" ><font color=#ffffff><b>Item Description</b></font></td>
          <td align="center" ><font color=#ffffff><b>Reqd Date</b></font></td>
          <td align="center" ><font color=#ffffff><b>Qty</b></font></td>
          <td align="center" ><font color=#ffffff><b>Possible Source</b></font></td>
          <td align="center" ><font color=#ffffff><b>Project</b></font></td>
          <td align="center" ><font color=#ffffff><b>Purpose</b></font></td>
          <td align="center" ><font color=#ffffff><b>Velankani Asset</b></font>
          </td>
          <td align="center" ><font color=#ffffff><b>Purchase/Service</b></font></td>
          <td align="center" ><font color=#ffffff><b>Status</b></font></td>
        </tr>
        <%
	   Dim ReqID,i,PrjID,sPrjName,sPur_Ser,sCurr,iCount
	   Dim QtyReq,QtyApp
	   lclstr_bgColor = gsBGColorLight
		ReqID = Request.Form("hdReqId")
		sql = sql_GetPurchaseRequestItems(ReqID)
		call RunSql(sql,rsItems)
		i = 1
		Do while rsItems.EOF = false
			PrjId = rsItems("ProjectId")
			sql = sql_GetProjectName(PrjID)
			call RunSql(sql,rsPrj)
			if not rsPrj.EOF then
				sPrjName = rsPrj("ProjectName")
			end if
			if rsItems("PurchaseOrService") = 0 then
				sPur_Ser = "Purchase"
			else
				sPur_Ser = "Service"
			end if
			QtyReq = (rsItems("QuantityRequested") - rsItems("QuantityApproved"))
			if rsItems("Status") = 0 then
				sAction = "New Approval"
			else
				sAction = "Partially Approved"
			end if


	%>
        <tr>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" nowrap>
            <input type="Checkbox" name="action_<%=i%>" value="<%=i%>" onClick="Check(this.value);">
            <%=i%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><%=rsItems("ItemDescription")%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><%=SetDateFormat(rsItems("RequiredDate"))%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" >
            <input type="text" class=formstyleTooShort name="Quantity_<%=i%>" maxlength="4" value="<%=QtyReq%>" style="border: 1 solid" onblur="javascript:validateQuantity(this);" onfocus="javascript:isToPropagate=true;" DISABLED >
            <input type="hidden" name="hdQty<%=i%>" value="<%=QtyReq%>">
          </td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><%=rsItems("PossibleSource")%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><%=sPrjName%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><%=rsItems("Purpose")%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" >
            <input type="checkbox" name="Ownby_<%=i%>" value="isChecked" DISABLED>
          </td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><%=sPur_ser%></td>
          <td align="center" bgcolor="<%=lclstr_bgColor%>" ><b><%=sAction%></b></td>
          <input type="hidden" name="hdPrjId_<%=i%>" value="<%=PrjID%>">
          <input type="hidden" name="hdReqID_<%=i%>" value="<%=ReqID%>">
          <input type="hidden" name="hdIDesc_<%=i%>" value="<%=rsItems("ItemDescription")%>">
        </tr>
        <%
		i = i+1
		rsItems.movenext
		Loop
		iCount=rsItems.recordcount
		rsItems.close
		%>
      </table>
    </td>
  </tr>
  <tr>
  <td>&nbsp;</td>
  </tr>
  <tr>
  <td>&nbsp;</td>
  </tr>
  <tr>
    <td bgcolor=#ffffff width=100%>
      <table width="20%" align="center" cellspacing=2 cellpadding=2 border=0>
        <tr height=25>
          <td bgcolor="<%=lclstr_bgColor%>" align=center>
            <input class=formbutton type=button value="Approve" name="Approve" style="border: 1 solid" onclick="submit_approve()">
          </td>
          <td width=60>&nbsp;</td>
          <td bgcolor="<%=lclstr_bgColor%>" align=center>
            <input class=formbutton type=button value="Reject" name="Reject" style="border: 1 solid" onclick="submit_reject();">
          </td>
		  <td><input type="hidden" name="AppItems" value=""></td>
		  <td><input type="hidden" name="RejItems" value=""></td>
	  	  <td><input type="hidden" name="hdCount" value="<%=iCount%>"></td>

        </tr>
      </table>
    </td>
  </tr>
  <tr bgcolor=#ffffff>
    <td>&nbsp;</td>
  </tr>

  <tr bgcolor=#ffffff>
    <td   align=center><A href="../../../main.asp"><%=dictLanguage("Return_Business_Console")%></a></td>
  </tr>
  <tr bgcolor=#ffffff>
    <td>&nbsp;</td>
  </tr>
</table>
</form>
<br>

<!--#include file="../includes/connection_close.asp"-->
