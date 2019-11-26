

<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/Connection_open.asp"-->
<!--#include file="../includes/style.asp"-->


<%
Dim ReqID,sReqNo,sEName,sEId 
ReqID =  Request.QueryString("lID")
sReqNo = Request.QueryString("sReq")
sDept = Request.QueryString("sDept")
sEName = Request.QueryString("sEmpName")
sEId = Request.QueryString("sEmpID")
%><title><%=gsSiteName%></title>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    <td colspan="3" align="right"><img src="<%=gsSiteRoot%>gif/logo/Velankani_logo.gif" border = "0" ></td>
  </tr>


  <tr> 
    <td align="center" colspan="3"><font size="5"><b>Purchase Request</b></font></td>
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
    <td align="Left" colspan="3">
	  <table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor="#666666" border="0" >
        <tr > 
          <td align="left"  width="25%" bgcolor="#FFFFFF"><p style="margin-left:5"><b>DEPARTMENT</b></p> </td>
          <td width="25%" bgcolor="#FFFFFF"><p style="margin-left:10"><%=sDept%></p></td>
          <td align="left"  width="25%" bgcolor="#FFFFFF"><p style="margin-left:5"><b>PURCHASE REQUEST 
            NUMBER</b></p></td>
          <td  width="25%" bgcolor="#FFFFFF" ><p style="margin-left:10"><%=sReqNo%></p></td>
        </tr>
      </table>
	</td>
  </tr>
  <tr> 
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td align="Left" colspan="3">
	  <table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor="#666666" border="0">
        <tr > 
          <td align="center" bgcolor="#FFFFFF"><b>Sl No.</b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>ITEM DESCRIPTION</b> 
          </td>
          <td align="center" bgcolor="#FFFFFF"><b>QTY REQUIRED</b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>REQUIRED DATE</b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>APPROX COST</b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>POSSIBLE SOURCE</b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>PROJECT </b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>PURPOSE</b> </td>
          <td align="center" bgcolor="#FFFFFF"><b>SPECIAL INSTRUCTIONS</b> </td>
        </tr>
           <%
			if ReqID <> "" then		   
				sql="sp_PSystem_GetItemsByPurchaseRequisitionId " & ReqID & " "
				call RunSql(sql,rsItems)
				i=1
				while not rsItems.eof
			%>

        <tr > 
          <td align="center" bgcolor="#FFFFFF"><%=i%></td>
          <td align="center" bgcolor="#FFFFFF"> <%=rsItems("ItemDescription")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("QuantityRequested")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("RequiredDate")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("ApproxUnitCost")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("PossibleSource")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("Project")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("Purpose")%></td>
          <td align="center" bgcolor="#FFFFFF"><%=rsItems("SpecialInstruction")%></td>
        </tr>
		<%
			i =i  + 1
			rsItems.movenext	
			Wend
			rsItems.Close
	     end if	
		%>
      </table>
	
	</td>
  </tr>
  <tr> 
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
  <tr> 
    <td align="Left" colspan="3">
	  <table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor="#666666" border="0">
        <tr> 
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Requisitioner's Name:</b></p>
          </td>
          <td  width="17%"  bgcolor="#FFFFFF" align="center"><%=sEName%> ( <%=sEId%> )</td>
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Director / Project Manager:</b></p>
          </td>
          <td  width="17%" bgcolor="#FFFFFF" align="center">&nbsp;</td>
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>VP (O), Authorized Signatory:</b></p>
          </td>
          <td  width="17%" bgcolor="#FFFFFF" align="center">&nbsp;</td>
        </tr>
        <tr> 
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Signature:</b></p>
          </td>
          <td bgcolor="#FFFFFF" width="17%">&nbsp;</td>
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Signature:</b></p>
          </td>
          <td bgcolor="#FFFFFF" width="17%">&nbsp;</td>
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Signature:</b></p>
          </td>
          <td bgcolor="#FFFFFF" width="17%">&nbsp;</td>
        </tr>
        <tr> 
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Date:</b></p>
          </td>
          <td  width="17%" bgcolor="#FFFFFF" align="center"><%=SetDateFormat(Date())%></td>
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Date:</b></p>
          </td>
          <td bgcolor="#FFFFFF" width="17%">&nbsp;</td>
          <td align="left" bgcolor="#FFFFFF" width="17%"> 
            <p style="margin-left:5"><b>Date:</b></p>
          </td>
          <td bgcolor="#FFFFFF" width="17%">&nbsp;</td>
        </tr>
      </table>
	
	</td>
  </tr>
  <tr> 
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
    <tr> 
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>
    <tr> 
    <td align="Left" colspan="3">
	  <table width="90%" cellpadding="1" cellspacing="1" align="center" bgcolor="#666666" border="0">
        <tr height="20"> 
          <td align="left" bgcolor="#FFFFFF" ><b>Special instructions regarding the order (if any): </b></td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF">
		  <br><br><br>
		  </td>
        </tr>
      </table>
	</td>
  </tr>
      <tr> 
    <td align="Left" colspan="3">&nbsp;</td>
  </tr>

    <tr> 
    <td align="Right"> 
      <INPUT  type="button" value="Print" class=formbutton  style="border: 1 solid" name=button2 onclick="window.print();">
    </td>
    <td align="Right">&nbsp;</td>
    <td align="Left"> 
      <input type="button" value="Close" class=formbutton  style="border: 1 solid" name=Close onclick="self.close();">
    </td>
  </tr>
</table>



<!--#include file="../includes/connection_close.asp"-->