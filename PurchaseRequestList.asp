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
          <td align="center" width="95%"><font color="#ffffff"><b>Approver - Purchase Requests Inbox</b></font></td>
          <td align="right" width="5%"><p style="margin-right:10"><a href="ApproverInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></p></td>
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
                <td align="center" ><font color=#ffffff><b>Requisition No</b></font></td>
                <td align="center" ><font color=#ffffff><b>Request Date</b></font></td>
                <td align="center" ><font color=#ffffff><b>Status</b></font></td>
              </tr>
              <%
			  	Dim sReqID,str,iCounter

				lclstr_bgColor = gsBGColorLight
				sql = sql_GetPurchaseRequestList()

				call RunSql(sql,rsList)
				iCounter = 1
				if  rsList.EOF = false then

				While NOT rsList.Eof
					sReqId = rsList("RequisitionID")
					sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& sReqId &" "
					call RunSql(sql,rsRec)
					if rsRec.Eof = false then
						ReqNum = rsRec("RequisitionNum")
					end if
					rsRec.close

					'if rsList("Status") = 4 then
					'	sStatus = "Quotations to approve"
					'else
					'	sStatus = "PR for approval"
					'end if
					if lclstr_bgColor = gsBGColorLight then
						lclstr_bgColor = gsBGColorDark
					else
						lclstr_bgColor = gsBGColorLight
					end if
			  %>
              <tr bgcolor="<%=lclstr_bgColor%>">
                <td align="center"> <%=iCounter%> </td>
                <td align="center"> <a href="javascript:redirect(<%=sReqId%>)"><%=GetPurchaseRequisitionNo(ReqNum)%></a>
                </td>
                <td align="center"><%=SetDateFormat(rsList("Requisitiondate"))%></td>
                <td align="center"><b><%=rsList("Status")%></b></td>
              </tr>
              <%
				iCounter = iCounter + 1
				rsList.movenext
			Wend
			else
				Response.Write "<tr ><td align='center' colspan='3'><br><b><font color='red'> No records found </font></b></td></tr>"
			end if
				rsList.close
			  %>
              <tr class="blue">
                <td align="center" colspan="4">
                  <div align="right">&nbsp;</div>
                </td>
              </tr>
            </table>




          </td>
        </tr>
      </table>
<br>
 <script language="javascript">
 function redirect(requestId)
 {
 	document.FinalForm.hdReqId.value = requestId;

	document.FinalForm.method="Post";
	document.FinalForm.action="ApproverPR_View.asp"
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