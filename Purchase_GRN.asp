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


<Script language=JavaScript src="../includes/javascript/validate.js" type=text/javascript></SCRIPT>
<%
	ReceivedEditAction="Add"
	EmployeeId=Session("Employee_Id")

	sql="select EmployeeId from tbl_WKI_Admin where EmployeeId='" & EmployeeId  & "'"
	call RunSQL(sql,rsAdmin)
	if not rsAdmin.eof then
		isAdmin=true
	else
		isAdmin=false
	end if
	rsAdmin.close()

	sql="select DeptName from VSPL_Managers A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsDepartment)
	if not rsDepartment.eof then
		isManager=true
		if rsDepartment(0)="HR" then
			isHRManager=true
		elseif rsDepartment(0)="ITIT" then
			isITITManager=true
		else
			isHRManager=false
			isITITManager=false
		end if
	else
		isManager=false
	end if
	rsDepartment.close()

	sql="select DeptName from VSPL_EmployeeHREntry A,sql_DepartmentMaster B where A.DeptId=B.DeptId and A.EmployeeId='" & EmployeeId & "'"
	call RunSQL(sql,rsDepartment)
	if not rsDepartment.eof then
		if rsDepartment(0)="HR" then
			isHRMember=true
		elseif rsDepartment(0)="ITIT" then
			isITITMember=true
		else
			isHRMember=false
			isITITMember=false
		end if
	else
		isManager=false
	end if
	rsDepartment.close()
	PurRequisitionNo = Request.QueryString("PurRequisitionNo")

	Dim PurOrdNo
	PurOrdNo = Request.Form("PurOrdNo")
	

%>
<Script language="Javascript">
	function ItemInfo()
	{
		document.strFormm.method="post";
		document.strFormm.action="Purchase_GRN.asp";
		document.strFormm.submit();
	}
	function GetAddress(Address)
	{
		document.strFormm.SupAddress.value = Address;
		for (var i=0; i<document.strFormm.Supplier.length;i++)
		{
			if(document.strFormm.Supplier.options[i].selected== true)
			{
			var hide =(document.strFormm.Supplier.options[i].text);
			document.strFormm.hdSupName.value = hide;
			}
		}
	}
</Script>
<script language="javascript" type="text/javascript">
	var isToPropagate=true;
	var index=2;
	var editIndex=-1;

	function callOpenCalendar(ctrl)
	{
		if(ctrl.value=="")
		openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','PartyChallanDate',150,300);
	}
	function callOpenCalendar_1(ctrl)
	{
		if(ctrl.value=="")
		openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','DeliveryDate',150,300);
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

	function validatePartyChallanDate(ctrl)
	{
		var isToPropagate=true;
		var regexp=new RegExp(/(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		if(isToPropagate==true)
		{
			if(ctrl.value=="")
			{
				alert("Party Challan Date required field");
				isToPropagate=false;
				openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','PartyChallanDate',150,300);
				return false;
			}
			else if(!regexp.test(ChangeToMMDDYYYY(ctrl.value)))
			{
				alert("Please enter a valid Party Challan Date");
				isToPropagate=false;
				openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','PartyChallanDate',150,300);
				return false;
			}
			else
				isToPropagate=true;
		}
			return true;
	}

	function validateDeliveryDate(ctrl)
	{
		var isToPropagate=true;
		var regexp=new RegExp(/(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)\d\d/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		if(isToPropagate==true)
		{
			if(ctrl.value=="")
			{
				alert("Delivery Date is required field");
				isToPropagate=false;
				ctrl.focus();
			//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','DeliveryDate',150,300);
				return false;
			}
			else if(!regexp.test(ChangeToMMDDYYYY(ctrl.value)))
			{
				alert("Please enter a valid Delivery Date.");
				isToPropagate=false;
				ctrl.focus();
			//	openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','DeliveryDate',150,300);
				return false;
			}
			else if (isGreaterDate(ChangeToMMDDYYYY(document.strFormm.DeliveryDate.value),ChangeToMMDDYYYY(document.strFormm.PartyChallanDate.value)))
			{
				alert("Delivary Date cannot be less that PartyCahllan Date");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;
		}
			return true;
	}

	function validatePurOrderNo(ctrl)
	{
		if(isToPropagate==true)
		{
			if(ctrl.value=="0")
			{
				alert("Please select Purchase Order No");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;
		}
		return true;
	}
	function validatePartyChallanNo(ctrl)
	{
		var regexp = new RegExp (/[0-9a-zA-Z]/);
		ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
		if (isToPropagate==true)
		{
			if (ctrl.value=="")
			{
				alert("Party Challan No. is required field");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else if(!regexp.test(ctrl.value))
			{
				alert("Please enter valid Party challan no.");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			isToPropagate= true;
		}
		return true;
	}
	function validateSecurityEntryNo(ctrl)
	{
	var regexp = new RegExp (/[0-9a-zA-Z]/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(ctrl.value!="")
	{
		if (isToPropagate==true)	
		{
			if(!regexp.test(ctrl.value))
			{
				alert("Please enter Alpha-Numeric characters only.");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate = true;
		}		
	}
		return true;
	}
	function validateVehicleNo(ctrl)
	{
	var regexp = new RegExp (/[0-9a-zA-Z]/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(ctrl.value!="")
	{
		if (isToPropagate==true)
		{
			if(!regexp.test(ctrl.value))
			{
				alert("Please enter Alpha-Numeric characters only.");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;
		}
	}	
	return true;
	}

	function validateLLRRNo(ctrl)
	{
	var regexp = new RegExp (/[0-9a-zA-Z]/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if (isToPropagate==true)
	{
		if(ctrl.value == "")
		{
			alert("LL/RR no is required field");
			isToPropagate=false;
			ctrl.focus();
			return false;
		}
		else if(!regexp.test(ctrl.value))
		{
			alert("Please enter Alpha-Numeric characters only.");
			isToPropagate=false;
			ctrl.focus();
			return false;
		}
		else
			isToPropagate=true;
	}
	else
		return true;
	}
	function validateSupplier(ctrl)
	{
	if (isToPropagate==true)
	{
		if(ctrl.value=="0")
		{
			alert("Select Supplier");
			isToPropagate=false;
			ctrl.focus();
			return false;
		}
		else
			isToPropagate=true;
	}	
	return true;
	}

	function validateDescription(ctrl)
	{
	var regexp = new RegExp (/[0-9a-zA-Z]/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(isToPropagate==true)
	{
		if(ctrl.value!="")
		{
			if(!regexp.test(ctrl.value))
			{
				alert("Please enter Alpha-Numeric characters only.");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
				isToPropagate=true;		
		}
	}
	return true;
	}
	function validateRemarks(ctrl)
	{
	var regexp = new RegExp (/[0-9a-zA-Z]/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(ctrl.value!="")
	{
		if (isToPropagate==true)
		{
			if(!regexp.test(ctrl.value))
			{
				alert("Please enter Alpha-Numeric characters only.");
				isToPropagate=false;
				ctrl.focus();
				return false;
			}
			else
			isToPropagate=true;
		}	
	}
	return true;
	}

	function ValidateQtyAcc(ctrl)
	{
		var a = ctrl.name;
		var indexarr = a.split("_");
		var Num = indexarr[1];
		var Rec = document.getElementById("QtyReceived_" + Num);
		var Acc = document.getElementById("QtyAccepted_" + Num);
		var TillDate = document.getElementById("QtyTillDate_" + Num);
		var QtyTD  = TillDate.value;
		
		var x = Rec.value;
		if (x == "")		
		{
			alert("Please enter quantity received");
			document.getElementById("QtyReceived_" + Num).focus();
			return false;
		}
		function significantnumber( x, significance ) 
		{
		  if(typeof significance==='undefined') { significance = 2; }
		  x = Math.round( x * Math.pow( 10, significance ) ) / Math.pow( 10, significance );
		  return x; // may need padding
		}
		var y = Acc.value;		
		function significantnumber( y, significance ) 
		{
		  if(typeof significance==='undefined') { significance = 2; }
		  y = Math.round( y * Math.pow( 10, significance ) ) / Math.pow( 10, significance );
		  return y; // may need padding
		}
		if (parseInt(y) > parseInt(x))
		{
			alert("Quantity accepted can't be more than quantity received");
			document.getElementById("QtyAccepted_" + Num).value= "0";
			document.getElementById("QtyAccepted_" + Num).focus();
			return false;
		}
		else if (parseInt(y) > parseInt(QtyTD))
		{
			alert("Quantity accepted can't be more than" + " " + QtyTD);
			document.getElementById("QtyAccepted_" + Num).value= "0";
			document.getElementById("QtyAccepted_" + Num).focus();
			return false;
			
		}
		else
			var z = x - y;
			document.getElementById("QtyRejected_" + Num).value = z;
			
	}
	function ValidateQtyRec(ctrl)
	{

	var a = ctrl.name;
	var indexarr = a.split("_");
	var Num = indexarr[1];
	var Rec = document.getElementById("QtyReceived_" + Num);
	var x = Rec.value;
/*	if (x == "")
	{
		alert("Quantity received is required field");
		document.getElementById("QtyReceived_" + Num).focus();
		return false;
	}
*/	
	var Qty = document.getElementById("hdQty_" + Num);
	var Acc = document.getElementById("QtyAccepted_" + Num);
	var y = Acc.value;
		function significantnumber( x, significance ) 
		{
		  if(typeof significance==='undefined') { significance = 2; }
		  x = Math.round( x * Math.pow( 10, significance ) ) / Math.pow( 10, significance );
		  return x; // may need padding
		}
		//alert (parseInt(x));
		if (parseInt(x) > (Qty.value))
		{
			alert("Quanitity received can't be more than" + " " + Qty.value);	
			document.getElementById("QtyReceived_" + Num).value="0";
			document.getElementById("QtyReceived_" + Num).focus();
			return false;
		}
		else if (parseInt(y) > parseInt(x))
		{
			alert("Quantity accepted can't be more than quantity received");
			document.getElementById("QtyAccepted_" + Num).value= "0";
			document.getElementById("QtyAccepted_" + Num).focus();
			return false;
		}
		else
			var z = x - y;
			document.getElementById("QtyRejected_" + Num).value = z;
	}
	function EnableElements(k)
	{
		//alert(k);
		if (document.getElementById("action_"+k).checked == true)
		{
			document.getElementById("QtyReceived_"+k).disabled = false;
			document.getElementById("QtyAccepted_"+k).disabled = false;
		//	document.getElementById("QtyRejected_"+k).disabled = false;
		}
		else
		{
			document.getElementById("QtyReceived_"+k).disabled = true;
			document.getElementById("QtyAccepted_"+k).disabled = true;
		//	document.getElementById("QtyRejected_"+k).disabled = true;
			document.getElementById("QtyReceived_"+k).value = "";
			document.getElementById("QtyAccepted_"+k).value = "";
			document.getElementById("QtyRejected_"+k).value = "";
		}	
	}
	function validateQtyReceived(ctrl)
	{
	var regexp = new RegExp (/^[0-9]\d*$/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(isToPropagate==true)
	{
		if (ctrl.value!="")
		{
		 if(!regexp.test(ctrl.value))
			{
				alert("Please enter a valid Quantity Received.");
				isToPropagate=false;
				ctrl.value="0";
				ctrl.focus();
				return false;
			}
			else
			isToPropagate=true;
		}
		}
			return true;
	}
	function validateQtyAccepted(ctrl)
	{
	var regexp = new RegExp (/^[0-9]\d*$/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(isToPropagate==true)
	{
		if (ctrl.value!="")
		{
		 if(!regexp.test(ctrl.value))
			{
				alert("Please enter a valid Quantity Accepted.");
				isToPropagate=false;
				ctrl.value="";
				ctrl.focus();
				return false;
			}
			else
			isToPropagate=true;
		}
		}
		return true;
	}

	function validateQtyRejected(ctrl)
	{
	var regexp = new RegExp (/^[0-9]\d*$/);
	ctrl.value=ctrl.value.replace(/^\s+|\s+$|\b\s+(?=[^\w\s])|\s+(?=\s)/g,""); // replace multiple space with single space
	if(isToPropagate==true)
	{
	if(ctrl.value!=="")
	{
	 if(!regexp.test(ctrl.value))
		{
			alert("Please enter a valid Quantity Rejected.");
			isToPropagate=false;
			ctrl.value="";
			ctrl.focus();
			return false;
		}
		else
		isToPropagate=true;
	}
	}
		return true;
	}
function Validator(frm)
	{
		var str,s,i
    	formElements=["PartyChallanNo","PartyChallanDate","DeliveryDate","LLRRNo"];
     	for(i=0;i<3;i++)
    	{
	      if(frm.elements[formElements[i]].value.length !=0)
    	  {
        	 str=frm.elements[formElements[i]].value
	         s = str.replace(/^(\s)*/, '');
	         s = s.replace(/(\s)*$/, '');
	         frm.elements[formElements[i]].value=s
    	  }
	    }
		var x=document.strFormm.elements.length;
		var p=0;
		var half=0;
		var cnt=0;
		if(document.strFormm.PurOrdNo.value=="0")
			{
			alert("Select Purchase Order No.");
			document.strFormm.PurOrdNo.focus();
			return false;
			}
		for(i=0;i<x;i++)
		{
		if(document.strFormm.elements[i].type=="checkbox")
			{
			p=p+1;
	
			}
		}
		half=p;
		for(k=1;k<=half;k++)
		{
			if (document.getElementById("action_"+k).checked == false)
			  {
				cnt=cnt+1;
			  }
		} //for loop
	
		  if (half==cnt)
		  {
		   alert('Please select an item for GRN Entry');
		   return false;
		  }
		for(k=1;k<=half;k++)
		{
			if (document.getElementById("action_"+k).checked == true)
			  {
				if (document.getElementById("QtyReceived_" + k).value == "")
					{
						alert("Quantity received is required field");
						document.getElementById("QtyReceived_" + k).focus();
						return false;
					}
				else if (document.getElementById("QtyAccepted_" + k).value == "")	
					{
						alert("Quantity accepted is required field");
						document.getElementById("QtyAccepted_" + k).focus();
						return false;
					}
			  }
		} //for loop
		

		if(frm.PartyChallanNo.value == "")
			{
			alert("Party Challan No. is required field");
			frm.PartyChallanNo.focus();
			return false;
			}
		if (frm.PartyChallanDate.value == "")
			{
			alert("Party Challan Date is required field");
			frm.PartyChallanDate.focus();
			return false;
			}
		if (frm.DeliveryDate.value == "")
			{
			alert("Delivery Date is required field");
			frm.DeliveryDate.focus();
			return false;
			}
		if (frm.LLRRNo.value == "")
			{
			alert("LL/RR no is required field");
			frm.LLRRNo.focus();
			return false;
			}

		if (isGreaterDate(ChangeToMMDDYYYY(document.strFormm.DeliveryDate.value),ChangeToMMDDYYYY(document.strFormm.PartyChallanDate.value)))
		{
			alert("Delivary Date cannot be less that PartyCahllan Date");
			document.strFormm.DeliveryDate.focus();
			return false;
		}
		return true;
	}

</script>

      <table width="100%" cellspacing="2" cellpadding="2" border="0">
        <tr height="25" class="blue" align="center">
          <td align="center" width="95%"> <font color=#ffffff><b>GOODS RECEIVED NOTE (GRN)</b></font></td>
          <td align="right" width="5%"><p style="margin-right:10"><a href="PurchaseTeamInbox.asp" style="text-decoration:none"><font color="#ffffff">Inbox</font></a></p></td>
        </tr>
        <tr hight="25">
          <td colspan="2">&nbsp;</td>
        </tr>
        <tr height="25" align="left">
          <td colspan="2" >
            <form name="strFormm" method="Post" action="Submit_GRN.asp" onSubmit ='return Validator(this)'>
              <table width="100%" align="center" cellspacing="2" cellpadding="2">
                <tr class="blue"> 
                  <td align="left" colspan="9"><font color="#ffffff"><b>Purchase 
                    Order Released List</b></font></td>
                </tr>
                <tr bgcolor="<%=gsBGColorLight%>"> 
                  <td colspan="9"> 
                    <Select class="formstylemed" name="PurOrdNo" onChange="ItemInfo(this.value)" >
                      <option Selected value="0">Select Purchase Released List</option>
                      <%
					  	sql = " Select distinct tbl_Psystem_PurchaseOrder.PurOrderNo,tbl_Psystem_PurchaseOrder.PurOrderNum,tbl_Psystem_PurchaseOrder.RequisitionId from tbl_Psystem_PurchaseOrder inner join tbl_Psystem_Quotations " & _
							" ON (tbl_Psystem_PurchaseOrder.PurOrderNo = tbl_Psystem_Quotations.PurOrderNo) where (tbl_Psystem_Quotations.isGRNEntered = 0 " & _
							" or tbl_Psystem_Quotations.isGRNEntered = 2) and tbl_Psystem_Quotations.isApproved = 4 and tbl_Psystem_Quotations.isPOCancelled = 0 Order by tbl_Psystem_PurchaseOrder.PurOrderNum"
						
						Call RunSql(sql,rsList)
						
						While Not rsList.EOF
						PurOrderNo = rsList("PurOrderNo")
						PurOrderNum = rsList("PurOrderNum")
						if cInt(rsList("PurOrderNo")) = cInt(PurOrdNo) then
			 		%>
                      <option Selected value="<%=rsList("PurOrderNo")%>"><%=GetPurchaseOrderNo(PurOrderNum)%></option>
                      <% else %>
                      <option value="<%=rsList("PurOrderNo")%>"><%=GetPurchaseOrderNo(PurOrderNum)%></option>
                      <% end if %>
                      <%
					 	rsList.movenext
						Wend
						rsList.close
				  %>
                    </select>
                  </td>
                </tr>
                <tr class="blue"> 
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Sl.No</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Item Description</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Project</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Tax Percent</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Quantity</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Qty Received</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Qty Accepted</b></font></div>
                  </td>
                  <td> 
                    <div align="center"><font color="#ffffff"><b>Qty Rejected</b></font></div>
                  </td>
                  <td> 
                    
                  <div align="center"><font color="#ffffff"><b>Till Date Quantity 
                    Received</b></font></div>
                  </td>
				  
                </tr>
			<%
				if cInt(PurOrdNo) <> "" then
					sql="Select * from tbl_Psystem_Quotations where isApproved= 4 and (isGRNEntered = 0  or isGRNEntered = 2) and purOrderNo = " & cInt(PurOrdNo) &" "
					call RunSql(sql,rsPurOrd)
					i = 1
					While rsPurOrd.EOF = false
						ReqId  = rsPurOrd("RequisitionId")
						PrjId = rsPurOrd("ProjectId")
						sql = sql_GetProjectName(PrjID)
						call RunSql(sql,rsPrj)
						if not rsPrj.EOF then	
							sPrjName = rsPrj("ProjectName")
						end if
							rsPrj.Close						
						 Qty = rsPurOrd("Quantity")
						 TaxPercent = rsPurOrd("TaxPercent")
 						 sSupName =  rsPurOrd("SupplierName")
						 sql = "Select SupplierAddress from tbl_Psystem_Supplier where SupplierName = '" & sSupName & "' "
						 Call RunSql(sql,rsSup)
						 if Not rsSup.EOF then
							 sSupAddr = rsSup("SupplierAddress")
						 end if
						 rsSup.Close
						 
						
				%>
                <tr bgcolor="<%=gsBGColorLight%>"> 
                  <td> 
                    <div align="center">
					<input type="Checkbox" name="action_<%=i%>" value="<%=i%>" onClick="EnableElements(this.value);"><%=i%>
					</div>
                  </td>
                  <td> 
                    <div align="center"><%=rsPurOrd("ItemDescription")%></div>
                  </td>
                  <td> 
                    <div align="center"><%=sPrjName%></div>
                  </td>
                  <td> 
                    <div align="center"><%=rsPurOrd("TaxPercent") & " " & "%" %></div>
                  </td>
                  <td> 
                    <div align="center"><%=Qty%></div>
                  </td>
                  <td> 
                    <div align="center">
                      <input type="text" name="QtyReceived_<%=i%>" maxlength="4" size="6" onfocus="javascript:isToPropagate=true;" onblur="javascript:validateQtyReceived(this);ValidateQtyRec(this)" DISABLED>
                    </div>
					<input type="hidden" name="hdQty_<%=i%>" value="<%=Qty%>">
                  </td>
                  <td> 
                    <div align="center">
					<input type="text" name="QtyAccepted_<%=i%>" maxlength="4" size="6" onfocus="javascript:isToPropagate=true;" onblur="javascript:validateQtyAccepted(this);ValidateQtyAcc(this)" DISABLED>
                    </div>
					
                  </td>
                  <td> 
                    <div align="center">
					<input type="text" name="QtyRejected_<%=i%>" maxlength="4" size="6" onfocus="javascript:isToPropagate=true;" onblur="javascript:validateQtyRejected(this);" Readonly>
					</div>
                  </td>
				  <input type="hidden" name="hdItemDesc_<%=i%>" value="<%=rsPurOrd("ItemDescription")%>">
				  <input type="hidden" name="hdPrjID_<%=i%>" value="<%=PrjId%>">
				  <input type="hidden" name="hdSupplier_<%=i%>" value="<%=sSupName%>">
				  
                  <td> 
				  <%
				  sql = "Select  Sum(QtyAccepted) as QtyAccepted from tbl_Psystem_GRN where PurOrderNo = "& PurOrdNo &" and ItemDescription = '" & Replace(Server.HTMLEncode(rsPurOrd("ItemDescription")),"'","''") & "' and SupplierName = '" & sSupName & "' and (isAccepted = 0 or isAccepted = 1)"
				  Call RunSql(sql,rsQty)					  
				  
				if rsQty("QtyAccepted") <> "" then
					Qty_Rec = rsQty("QtyAccepted")
				else
					Qty_Rec = "0"
				end if	
				
				QtyTillDate = (cInt(Qty) - cInt(Qty_Rec))
			   rsQty.close	
				  %>
                    <div align="center">
					
					<%=Qty_Rec%>	
					</div>
					<input type="hidden" name="QtyTillDate_<%=i%>" value="<%=QtyTillDate%>">
                  </td>
				  
                </tr>
				<%
					i = i + 1
					rsPurOrd.movenext
					Wend
					iCount=rsPurOrd.recordcount
					rsPurOrd.Close
				end if
				%>
              </table>
           
          
        
        <tr>
          <td align="center" valign="top" colspan="2">
          
		  
              <table width="90%" align="center" id="OrderRelease">
                <tr align="center" > 
                  <td colspan="4"><b> GRN Entry Details </b></td>
                </tr>
                <tr align="center" > 
                  <td colspan="4"><font color="red"><b>*</b></font><b>&nbsp;Fields 
                    are mandatory</b></td>
                </tr>
                <tr> 
                  <td class="blue" align="right"><font color=#ffffff><b>Party 
                    Challan No&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>"> 
                    <input class="formstylemedium" type="text" size="25" name="PartyChallanNo" maxlength="50" onFocus="javascript:isToPropagate=true;">
                    &nbsp;&nbsp; </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Party 
                    challan Date&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>"> 
                    <input class="formstylemedium" type="text" name="PartyChallanDate" size="11" value=""  onfocus="javascript:isToPropagate=true;" Readonly>
                    &nbsp; &nbsp;<a name="CalanderLink" onClick="openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','PartyChallanDate',150,300)"><img border="0" src="/gif/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" width="16" height="15"></a>&nbsp;&nbsp; 
                  </td>
                </tr>
                <tr> 
                  <td class="blue" align="right"><font color=#ffffff><b>Security 
                    Gate Entry No : </b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>"> 
                    <input type="text" class=formstylemedium size="25" name="SecutiryEntry" value="" maxlength="20" onFocus="javascript:isToPropagate=true;">
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Date of 
                    Delivery&nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>"> 
                    <input class="formstylemedium" type="text" name="DeliveryDate" size="11" value=""  onfocus="javascript:isToPropagate=true;" Readonly>
                    &nbsp; &nbsp;<a name="CalanderLink" onClick="openCalendar1('<%=SetDateFormat(Formatdatetime(now(),2))%>','Date_Change','DeliveryDate',150,300)"><img border="0" src="/gif/calendaricon.jpg" onMouseOver="this.style.cursor='hand'" width="16" height="15"></a>&nbsp;&nbsp; 
                  </td>
                </tr>
                <tr> 
                  <td class="blue" align="right"><font color=#ffffff><b>LL/RR 
                    No &nbsp;:<font color="red">*</font></b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>"> 
                    <input class="formstylemedium" type="text" size="25" name="LLRRNo" maxlength="50" onFocus="javascript:isToPropagate=true;">
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Vehicle 
                    No &nbsp;: </b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>"> 
                    <input class="formstylemedium" type="text" size="25" name="VehicleNo" maxlength="15" onFocus="javascript:isToPropagate=true;">
                  </td>
                </tr>
                <tr> 
                  <td class="blue" align="right"><font color=#ffffff><b>Supplier 
                    Name &nbsp;: </b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>" vAlign="middle" style="word-break: break-all; width:200px;"> <%=sSupName%> 
                  </td>
                  <td class="blue" align="right"><font color=#ffffff><b>Supplier 
                    Address : </b></font></td>
                  <td bgcolor="<%=gsBGColorLight%>" style="word-break: break-all; width:200px;"><%=sSupAddr%></td>
                </tr>
                <tr> 
                  <td class="blue" align="right" NOWRAP><font color="#ffffff"><b>Inspection 
                    results and remarks : </b></font></td>
                  <td colspan="3" bgcolor="<%=gsBGColorLight%>"> 
                    <textarea name="Remarks" rows="4" cols="20" onFocus="javascript:isToPropagate=true;"
					onKeyDown="textCounter(document.strFormm.Remarks,document.strFormm.remLen3,500)" onKeyUp="textCounter(document.strFormm.Remarks,document.strFormm.remLen3,500)"></textarea>
                    <input readonly type="hidden" name="remLen3" size="3" maxlength="3" value="500">
                    <font color="red" >Max Chars (500)</font></td>
                </tr>
                <tr> 
                  <td>&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                  <td >&nbsp;</td>
                </tr>
                <tr> 
                  <td  align="Center" colspan="4"> 
                    <input class="formbutton" type="Submit" name="AddButton" value="Submit" style="border: 1 solid;width:50Px" >
                    &nbsp; 
                    <input class="formbutton" type="Reset" name="ResetButton" value="Reset" style="border: 1 solid">
                    &nbsp; </td>
                </tr>
              </table>
              <input type="hidden" name="hdPurOrdNo" value="<%=PurOrdNo%>">
              <input type="hidden" name="hdReqId" value="<%=ReqId%>">
			 <input type="hidden" name="hdCount" value="<%=iCount%>">
			 
            </form>
          </td>
        </tr>
		
	    <tr>
          <td colspan="2">&nbsp;
					  
		   </td>
        </tr>

      </table>

<p align="center">

<SCRIPT LANGUAGE="JavaScript">
<!-- Web Site:  The JavaScript Source -->
<!-- Use one function for multiple text areas on a page -->
<!-- Limit the number of characters per textarea -->
<!-- Begin
function textCounter(field,cntfield,maxlimit) {
if (field.value.length > maxlimit) // if too long...trim it!
field.value = field.value.substring(0, maxlimit);

// otherwise, update 'characters left' counter
else
cntfield.value = maxlimit - field.value.length;
}
//  End -->
</script>
<p align="center">
<a href="../../main.asp"><%=dictLanguage("Return_Business_Console")%></a>
</p>

<!--#include file="../includes/main_page_close.asp"-->