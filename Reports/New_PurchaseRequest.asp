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
	FromDate = Request.form("txtFromDate")
	ToDate = Request.form("txtToDate")

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
	//alert(document.Export_PR.hdFromDate.value);
	//alert(document.Export_PR.hdToDate.value);
	document.Export_PR.method="Post";
	document.Export_PR.action="Report_New_PurchaseRequest.asp"
	document.Export_PR.submit();
	
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
    <td align="center" colspan="3"><font size="4"><b>Purchase Request</b></font></td>
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
	<%
	'action="Report_New_PurchaseRequest.asp"
	%>
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
              <input type="submit" name="Submit" value="Submit">&nbsp;&nbsp;
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
    <td  colspan="3"><%
	'Response.write FromDate & " " & ToDate
	if FromDate <> "" and ToDate <> "" then
	%>
	
	  <table width="90%" border="0" cellspacing="1" cellpadding="1" align="center" bgcolor="#999999">
        <tr bgcolor="#FFFFFF" valign="middle"> 
          <td colspan="10"> 
            <div align="center"> <% Response.write "Purchase Requests between <b>" &  FromDate & "</b> and <b>" & ToDate %></div>
          </td>
          <td>
            <div align="center">
			<input type="button" name="Export" value="Export-Excel" onClick="javascript:export_excel()">
			</div>
          </td>
        </tr>
        <tr bgcolor="#FFFFFF" valign="middle"> 
          <td> 
            <div align="center"><b>Sl.No</b></div>
          </td>
          <td > 
            <div align="center" ><b>Requisition No</b></div>
          </td>
          <td> 
            <div align="center"><b>Item Desctiption</b></div>
          </td>
          <td> 
            <div align="center"><b>Project</b></div>
          </td>
          <td> 
            <div align="center"><b>Quantity Requested</b></div>
          </td>
          <td> 
            <div align="center"><b>Possible Source</b></div>
          </td>
          <td> 
            <div align="center"><b>Special Instructions</b></div>
          </td>
          <td> 
            <div align="center"><b>Approx Unit Cost</b></div>
          </td>
          <td> 
            <div align="center"><b>Service Type</b></div>
          </td>
          <td> 
            <div align="center"><b>Requested Date</b></div>
          </td>
          <td> 
            <div align="center"><b>Requested Employee</b></div>
          </td>
        </tr>
        <%
	FromDate = Trim(Request.Form("txtFromDate"))
	ToDate = Trim(Request.Form("txtToDate"))
	
	sql = " Select distinct dbo.fn_Psystem_PurchaseRequisitionNo(b.RequisitionNum) as RequisitionID,a.ItemDescription,dbo.fn_TimeSheet_GetProjectName(a.ProjectId) as Project,a.QuantityRequested, " & _ 
		" a.PossibleSource,a.SpecialInstruction,a.ApproxUnitCost,dbo.fn_PSystem_isRupeeOrDollar(a.RupeeOrDollar) as Currency,dbo.fn_PSystem_isPurchaseOrService(a.PurchaseOrService) as ServiceType, " & _ 
		" dbo.fn_TSystem_GetVelankaniFormatDate(b.RequisitionDate) as RequestedDate,(dbo.fn_TSystem_EmployeeName(b.EmployeeID)+' ('+ b.EmployeeID+')') as Employee " & _ 
		" from tbl_Psystem_PurchaseRequestTransaction a,tbl_Psystem_PurchaseRequestMaster b where a.RequisitionId = b.RequisitionId and b.RequisitionDate " & _ 
		" between '" & FromDate &"' and '" & ToDate & "' "
		Call RunSql(sql,objRs)


		i = 1
		If Not objRs.EOF Then
			while Not objRs.EOF
			
	%>
        <tr bgcolor="#FFFFFF"> 
          <td> 
            <div align="center"><%=i%></div>
          </td>
          <td nowrap><%=objRs("RequisitionID")%></td>
          <td><%=objRs("ItemDescription")%></td>
          <td><%=objRs("Project")%></td>
          <td> 
            <div align="center"><%=objRs("QuantityRequested")%></div>
          </td>
          <td><%=objRs("PossibleSource")%></td>
          <td><%=objRs("SpecialInstruction")%></td>
          <td> 
            <div align="center"><%=objRs("ApproxUnitCost") & " " & objRs("Currency")%></div>
          </td>
          <td><%=objRs("ServiceType")%></td>
          <td> 
            <div align="center"><%=objRs("RequestedDate")%></div>
          </td>
          <td><%=objRs("Employee")%></td>
        </tr>
        <%
  	  i = i + 1	
	  objRs.movenext
	  wend
	 else
		Response. Write "<tr><td align='center' bgcolor='#ffffff' colspan='11'><b>No Records Found</b></td></tr>"
		end if
		set objRs = NOTHING
  %>
      </table>

	<%
	end if
	%>

	</td>
  </tr>
  <tr>
    <td  colspan="3">&nbsp;</td>
  </tr>
  <tr>
    <td  colspan="3">
	<form name="Export_PR">
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