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
	document.Export_PO.action="Report_GRN_Daily.asp"
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
    <td align="center" colspan="3"><font size="4"><b>Goods Received Note</b></font></td>
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
          <td align="Center" width="90%" bgcolor="#FFFFFF"> 
            <% Response.write "Goods Received Note between <b>" &  FromDate & "</b> and <b>" & ToDate %>
          </td>
          <td width="10%" bgcolor="#FFFFFF">
		  <input type="button" name="Export" value="Export-Excel" onClick="javascript:export_excel()">
		  </td>
        </tr>
        <tr> 
          <td bgcolor="#FFFFFF" colspan="2"> <br>
		  <%
		  	'sql = " Select distinct dbo.fn_Psystem_PurchaseOrderNo(PurOrderNo) as PurchaseOrderNo,PurOrderNo,dbo.fn_Psystem_GRNNo(GRNNo) as Goods,GRNNo,Remarks, " & _
			'	" dbo.fn_TSystem_GetVelankaniFormatDate(DeliveryDate) as DeliveryDate from tbl_Psystem_GRN " & _ 
			'	" Where DeliveryDate between '" & FromDate & "' and '" & ToDate & "' Order  by PurOrderNo "
			
			Sql = " Select distinct dbo.fn_Psystem_PurchaseOrderNo(a.PurOrderNum) as PurchaseOrderNo,a.PurOrderNo,b.GRNNo as GRNNo, " & _ 
				" b.Remarks as Remarks,dbo.fn_Psystem_GRNNo(b.GRNNum) as Goods,dbo.fn_TSystem_GetVelankaniFormatDate(b.Deliverydate) as Deliverydate from  " & _ 
				" tbl_Psystem_PurchaseOrder a,tbl_Psystem_GRN b where a.PurOrderNo = b.PurOrderNo and b.DeliveryDate " & _ 
				" between '"& FromDate &"' and '"& ToDate &"' Order  by a.PurOrderNo "
				
				Call Runsql(sql,objRs)
			if objRs.Eof = False then	
				Do While Not objRs.Eof
				
				PurOrdNo = objRs("PurOrderNo")
				DDate = objRs("DeliveryDate")
				Goods = objRs("Goods")
				GRNNo = objRs("GRNNo")
				Remarks = objRs("Remarks")
				
		
				
		  %>
            <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#666666">
              <tr bgcolor="#FFFFFF"> 
                <td > 
                  <p style="margin-left:10"><b>GRN No :</b></p>
                </td>
                <td> 
                  <p style="margin-left:10"><%=Goods%></p>
                </td>
                <td> 
                  <p style="margin-left:10"><b>Purchase Order No :</b></p>
                </td>
                <td> 
                  <p style="margin-left:10"><%=objRs("PurchaseOrderNo")%></p>
                </td>
                <td> 
                  <p style="margin-left:10"><b>Date :</b></p>
                </td>
                <td><p style="margin-left:10"><%=DDate%></p></td>
              </tr>
              <tr bgcolor="#FFFFFF" align="center"> 
                <td><b>Sl No</b></td>
                <td><b>Item Description</b></td>
                <td><b>Quantity Received</b></td>
                <td><b>Quantity Accepted</b></td>
                <td colspan="2"><b>Quantity Rejected</b></td>
              </tr>
              <%
				  sql = "Select ItemDescription,QtyReceived,QtyAccepted,QtyRejected from tbl_Psystem_GRN  " & _ 
				  	" where PurOrderNo = "& PurOrdNo &" and GRNNo = "& GRNNo &" "	
				  Call RunSql(sql,rsGRN)
				 ' Response.write sql
				  
				  i = 1
				  if rsGRN.EOF = false then
				  Do While Not rsGRN.EOF
					
			  %>
              <tr bgcolor="#FFFFFF" align="center"> 
                <td><%=i%></td>
                <td><%=rsGRN("ItemDescription")%></td>
                <td><%=rsGRN("QtyReceived")%></td>
                <td><%=rsGRN("QtyAccepted")%></td>
                <td colspan="2"><%=rsGRN("QtyRejected")%></td>
              </tr>
              <%
					i = i + 1
					rsGRN.movenext
					Loop
					end if
					rsGRN.Close
			  %>
              <tr bgcolor="#FFFFFF" > 
                <td> 
                  <p style="margin-left:10"><b>Remarks :</b></p>
                </td>
                <td colspan="5"><p style="margin-left:10"><%=Remarks%></p></td>
              </tr>
              <br>
            </table>
            <% 
				objRs.movenext
				Loop
				objRs.Close
			else
				Response. Write "<tr><td align='center' bgcolor='#ffffff' Colspan='2'><b>No Records Found</b></td></tr>"
			end if
			%>
			<br>
            <br>
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

	<%Response.Flush%>
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