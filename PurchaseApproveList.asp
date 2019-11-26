<%@LANGUAGE="VBSCRIPT"%>
<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Subhash Rampuri
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>

<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
	  <table width="100%" align="center" >
        <tr class="blue">
          <td align="center" width="95%"><font color="#ffffff"><b>Approved Purchase Request</b></font></td>
          <td align="right" width="5%"><p style="margin-right:10"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></p></td>
        </tr>
        <tr>
          <td align="center" colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td align="center" colspan="2">&nbsp;</td>
        </tr>
        <tr>
          <td align="center" colspan="2">
            <table width="50%" border="0" cellspacing="2" cellpadding="2">
              <tr class="blue">
                <td align="center" ><font color=#ffffff><b>Sl.No</b></font></td>
                <td align="center" ><font color=#ffffff><b>RequestID</b></font></td>
                <td align="center" ><font color=#ffffff><b>Status</b></font></td>
              </tr>
              <%
			  	Dim sReqID,str,iCounter

				lclstr_bgColor = gsBGColorLight
				sql = sql_GetPurchaseApprovalList()
				call RunSql(sql,rsList)
				iCounter = 1
				if rsList.Eof = fasle then
				While NOT rsList.EOF
					sReqId = rsList("RequisitionID")
					
					sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& sReqId &" "
					call RunSql(sql,rsRec)
					if rsRec.Eof = false then
						ReqNum = rsRec("RequisitionNum")
					end if
					rsRec.Close

					if lclstr_bgColor = gsBGColorLight then
						lclstr_bgColor = gsBGColorDark
					else
						lclstr_bgColor = gsBGColorLight
					end if
			  %>
              <tr bgcolor="<%=lclstr_bgColor%>">
                <td align="center"> <%=iCounter%></td>
                <td align="center"> <a href="Purchase_Quotation.asp?PurRequisitionNo=<%=sReqID%>"><%=GetPurchaseRequisitionNo(ReqNum)%></a>
                </td>
                <td align="center"><b>PR Approved</b></td>
              </tr>
              <%
				iCounter = iCounter + 1
				rsList.movenext
				Wend
				else
					Response.Write "<tr ><td align='center' colspan='3'><br><b><font color='red'> No records found </font></b></td></tr>"
				end if
				rsList.Close
			  %>
              <tr class="blue">
                <td align="center" colspan="3">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
<br>
<script language="javascript">
 function redirect(requistId)
 {
 	document.FinalForm.hdReqId.value=requistId;
	document.FinalForm.method="Post";
	document.FinalForm.action="Purchase_Quotation.asp"
	document.FinalForm.submit();

 }
</script>
<form name="FinalForm">
<input type="hidden" name="hdReqId" value="">
</form>
<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>
<!--#include file="../includes/main_page_close.asp"-->