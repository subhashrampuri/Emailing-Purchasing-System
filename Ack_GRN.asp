<%
'----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'iMorfus Intranet Systems - Version 4.0.0 ' - Copyright 2002 - 06 (c) i-Vista Digital Solutions Limited. All Rights Reserved.
'Usage of this software must meet the i-Vista Digital Solutions License terms and conditions. See the file iMorfuslicense.txt for more information.
'All Copyright notices must remain in place at all times.
'-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
'Developed By: Subhash rampuri
'_________________________________________________________________________________________________________________________________________________________________________________________________________________
%>
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<!--#include file="../includes/mail.asp"-->
<%
	'RequisitionId = Request.Form("hdReqId")

	EmployeeId=Session("Employee_Id")
	sql="select dbo.fn_PSystem_EmployeeDepartmentName('" & EmployeeId & "')"
	call RunSql(sql,rsEmployeeDepartmentName)
	EmployeeDepartmentName=rsEmployeeDepartmentName(0)
	rsEmployeeDepartmentName.close()
	sql="sp_PSystem_GetActiveApprover"
	call RunSql(sql,rsApprover)
	if rsApprover.eof then
		Response.write "<br><br><br><br><br><br><br><br><br><br><br><center><font color='red'><b>There is no member assigned in Approvers Panel.</b></font></center>"
		Response.end
	end if
	ApproverId=rsApprover("EmployeeId")
	ApproverEmail=rsApprover("EmployeeEmail")
	rsApprover.close()
	sql="sp_PSystem_GetLoggedEmployeeNameAndEmail '" & EmployeeId & "'"
	call RunSql(sql,rsEmp)
	EmployeeName=rsEmp("EmployeeName")
	EmployeeEmail=rsEmp("EmployeeEmail")
	rsEmp.close()
	
	GRNNo=Request.Form("hdGRN")
	
%>

<script language="JavaScript" type="text/javascript">
	function CallReqID()
	{
		myPopup = window.open('GRN_Print.asp?GRNNo=<%=GRNNo%>',42,'toolbar=no,width=800,height=600,scrollbars=yes,left=100,top=150,resizable=yes');
		if (!myPopup.opener)
	    myPopup.opener = self;
	}
</script>
      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr height="25" class="blue" align="center">
          <td align="center"> <b><font color="#ffffff">Goods Receievd Note</font></b></td>
        </tr>
        <tr height="25"  align="Center">
          <td align="Right"><a href="javascript:CallReqID();"><img src="images/printer.gif" width="20" height="21" border="0"></a></td>
        </tr>
        <tr>
          <td vAlign="top">
            <table width="90%" border="0" cellspacing="2" cellpadding="2" align="center">
              <%
		  	if GRNNo <> "" then
		  	sql = "Select distinct GRNNo,PurOrderNo,RequisitionId,PartyChallanNo,PartyChallanDate,SecurityEntryNo,DeliveryDate,LLRRNo,vehicleNo,SupplierName,Remarks from tbl_PSystem_GRN where GRNNo = "& GRNNo &" "
			Call RunSql(sql,rsGRN)			

			if rsGRN.EOF = false then
				GRNNo = rsGRN("GRNNo")
				sql = " Select GRNNum  from tbl_Psystem_GRN where GRNNo = "& GRNNo &" "
				call RunSql(sql,rsGRNNum)
				if rsGRNNum.Eof = false then
					GRNNum = rsGRNNum("GRNNum")
				end if
				rsGRNNum.Close

				PurOrderNo = rsGRN("PurOrderNo")
				sql = " Select PurOrderNum  from tbl_Psystem_PurchaseOrder where PurOrderNo = "& PurOrderNo &" "
				call RunSql(sql,rsPONum)
				if rsPONum.Eof = false then
					PurOrderNum = rsPONum("PurOrderNum")
				end if
				rsPONum.Close
				
				RequisitionId = rsGRN("RequisitionId")
				sql = " Select RequisitionNum  from tbl_Psystem_PurchaseRequestMaster where RequisitionId = "& RequisitionId &" "
				call RunSql(sql,rsRec)
				if rsRec.Eof = false then
					ReqNum = rsRec("RequisitionNum")
				end if
				rsRec.Close

				sSupName = rsGRN("SupplierName")
				PartyChallanNo = rsGRN("PartyChallanNo")
				PartyChallanDate = SetDateFormat(rsGRN("PartyChallanDate"))
				SecurityEntryNo = rsGRN("SecurityEntryNo")
				DeliveryDate = SetDateFormat(rsGRN("DeliveryDate"))
				LLRRNo = rsGRN("LLRRNo")
				VehicleNo = rsGRN("VehicleNo")
				Remarks = rsGRN("Remarks")
				sql= "Select SupplierAddress from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
				Call RunSql(sql,rsSup)
				if Not rsSup.Eof then
					sSupAddr = rsSup("SupplierAddress")
				end if
				rsSup.Close
		  %>
              <tr class="blue"> 
                <td> 
                  <div align="center"><font color="#ffffff"><b>GRN No </b></font></div>
                </td>
                <td> 
                  <div align="center"><font color="#ffffff"><b>Purchase Order 
                    No</b></font></div>
                </td>
                <td> 
                  <div align="center"><font color="#ffffff"><b>Purchase Request 
                    No</b></font></div>
                </td>
              </tr>
              <tr bgcolor="<%=gsBGColorLight%>"> 
                <td> 
                  <div align="center"><%=GetGRNNo(GRNNum)%></div>
                </td>
                <td> 
                  <div align="center"><%=GetPurchaseOrderNo(PurOrderNum)%></div>
                </td>
                <td> 
                  <div align="center"><%=GetPurchaseRequisitionNo(ReqNum)%></div>
                </td>
              </tr>
              <tr class="blue"> 
                <td> 
                  <div align="center"><font color="#ffffff"><b>Supplier Info </b></font></div>
                </td>
                <td> 
                  <div align="center"><font color="#ffffff"><b>Party Challan No</b></font></div>
                </td>
                <td> 
                  <div align="center"><font color="#ffffff"><b>Party Challan Date</b></font></div>
                </td>
              </tr>
              <tr> 
                <td vAlign="top" bgcolor="<%=gsBGColorLight%>" > 
                  <div align="center"></div>
                  <div align="center" style="word-break: break-all; width:250px;"> 
                      <% = sSupName  %>
                  </div>
                </td>
                <td bgcolor="<%=gsBGColorLight%>"> 
                  <div align="center"><%=PartyChallanNo%></div>
                </td>
                <td bgcolor="<%=gsBGColorLight%>"> 
                  <div align="center"><%=PartyChallanDate%></div>
                </td>
              </tr>
              <tr> 
                <td vAlign="top" bgcolor="<%=gsBGColorLight%>" align="center" rowspan="4" style="word-break: break-all; width:350px;">
				<%=sSupAddr%> 
                </td>
                <td class="blue"> 
                  <div align="center"><font color="#ffffff"><b>Security Gate Entry 
                    No </b></font></div>
                </td>
                <td class="blue"> 
                  <div align="center"><font color="#ffffff"><b>Delivery Date</b></font></div>
                </td>
              </tr>
              <tr> 
                <td bgcolor="<%=gsBGColorLight%>"> 
                  <div align="center"><%=SecurityEntryNo%></div>
                </td>
                <td bgcolor="<%=gsBGColorLight%>"> 
                  <div align="center"><%=DeliveryDate%></div>
                </td>
              </tr>
              <tr> 
                <td class="blue"> 
                  <div align="center"><font color="#ffffff"><b>LL/ RR No</b></font></div>
                </td>
                <td class="blue"> 
                  <div align="center"><font color="#ffffff"><b>Vehicle No</b></font></div>
                </td>
              </tr>
              <tr> 
                <td bgcolor="<%=gsBGColorLight%>"> 
                  <div align="center"><%=LLRRNo%></div>
                </td>
                <td bgcolor="<%=gsBGColorLight%>"> 
                  <div align="center"><%=VehicleNo%></div>
                </td>
              </tr>
              <%	
				end if	
				end if 
				%>
              <tr> 
                <td colspan="3" vAlign="top"> 
                  <table width="100%" border="0" cellpadding="2" cellspacing="2">
                    <tr class="blue"> 
                      <td> 
                        <div align="center"><font color="#ffffff"><b>Sl.No</b></font></div>
                      </td>
                      <td> 
                        <div align="center"><font color="#ffffff"><b>Item Description</b></font></div>
                      </td>
                      <td> 
                        <div align="center"><font color="#ffffff"><b>Quantity 
                          Received</b></font></div>
                      </td>
                      <td> 
                        <div align="center"><font color="#ffffff"><b>Quantity 
                          Accepted</b></font></div>
                      </td>
                      <td> 
                        <div align="center"><font color="#ffffff"><b>Quantity 
                          Rejected</b></font></div>
                      </td>
                    </tr>
                    <%
					sql = "Select ItemDescription,QtyReceived,QtyAccepted,QtyRejected from tbl_Psystem_GRN where GRNNo = "& GRnNo &" "
					Call runSql(sql,rsInfo)

					i = 1
					While Not rsInfo.Eof
					
					%>
                    <tr bgcolor="<%=gsBGColorLight%>" > 
                      <td> 
                        <div align="center"><%=i%></div>
                      </td>
                      <td> 
                        <div align="center"><%=rsInfo("ItemDescription")%></div>
                      </td>
                      <td> 
                        <div align="center"><%=rsInfo("QtyReceived")%></div>
                      </td>
                      <td> 
                        <div align="center"><%=rsInfo("QtyAccepted")%></div>
                      </td>
                      <td> 
                        <div align="center"><%=rsInfo("QtyRejected")%></div>
                      </td>
                    </tr>
                    <%
					i = i + 1
					rsInfo.movenext
					Wend
					rsInfo.Close

					%>
                  </table>
                </td>
              </tr>
              <tr> 
                <td colspan="3" bgcolor="<%=gsBGColorLight%>"><b>Remarks : </b> 
                  <%=Remarks%></td>
              </tr>
              <tr> 
                <td colspan="3">&nbsp;</td>
              </tr>
            </table>
          </td>
        </tr>
      </table>


<p align="center">
<a href="Purchase_GRN.asp" style="text-decoration:none"><b>Back</b></a> <br><br> <a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->