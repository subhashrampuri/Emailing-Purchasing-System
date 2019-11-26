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
<!--#include file="../includes/mail.asp"-->
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->
<%

	Dim PurRequisitionNo,iCount
	arrItems=split(Request.Form("ItemList"),",")
	iCount = 1
	For i = 0 to  ubound(arrItems)
		'Response.write arrItems(i) & "<br>"

	If iCounter = 12 Or i = UBound(arrItems) Then
			iCounter = 0
			ItemCode = "VSPL-" & Year(Date()) & Month(Date()) & Day(Date())& Hour(Time()) & Minute(Time()) & Second(Time()) &iCount

			ProjectId = Server.HTMLEncode(Replace(Replace(arrItems(i-12),"&#44;",","),"'","''"))
			ItemDescription = Server.HTMLEncode(Replace(Replace(arrItems(i-11),"&#44;",","),"'","''"))
			RecNo = Server.HTMLEncode(Replace(Replace(arrItems(i-10),"&#44;",","),"'","''"))
			vArray = Split(RecNo,"/")
			ReqNo = vArray(3)

			sql = " Select RequisitionId  from tbl_Psystem_PurchaseRequestMaster where RequisitionNum = "& ReqNo &" "
			call RunSql(sql,rsRec)
			if rsRec.Eof = false then
				PurRequisitionNo = rsRec("RequisitionId")
			end if
			rsRec.Close
			
			SupplierName = Server.HTMLEncode(Replace(Replace(arrItems(i-9),"&#44;",","),"'","''"))
			Price = Server.HTMLEncode(Replace(Replace(arrItems(i-8),"&#44;",","),"'","''"))
			if Server.HTMLEncode(Replace(Replace(arrItems(i-7),"&#44;",","),"'","''")) = "Rupee(s)" then
				Curr = 1
			else
				Curr = 0
			end if
			if arrItems(i-6) = "Inclusive" then
				'Tax = 1
			else
				'Tax = 0
			end if

			TaxPercent = Server.HTMLEncode(Replace(Replace(arrItems(i-5),"&#44;",","),"'","''"))
			if TaxPercent = "0" then
				Tax = 0
			else
				Tax = 1
			end if

			DeliveryTime = Server.HTMLEncode(Replace(Replace(arrItems(i-4),"&#44;",","),"'","''"))
			PaymentTerms = Server.HTMLEncode(Replace(Replace(arrItems(i-1),"&#44;",","),"'","''"))
			Warranty = Server.HTMLEncode(Replace(Replace(arrItems(i-3),"&#44;",","),"'","''"))
			Quantity = Server.HTMLEncode(Replace(Replace(arrItems(i-2),"&#44;",","),"'","''"))
			Remarks = Server.HTMLEncode(Replace(Replace(arrItems(i),"&#44;",","),"'","''"))


		sql =" sp_itbl_PSystem_Quotations " & PurRequisitionNo & ", '" & ItemDescription & "', '" & ItemCode & "' ," & ProjectId &", '" & SupplierName & "', " & Price & "," & Curr & ", " & Tax & ", " & TaxPercent & ", " & Quantity & ", '" & Warranty & "','" & DeliveryTime & "','" & PaymentTerms & "','" & Remarks & "' "
		'Response.write sql
		Call DoSql(sql)

		Else
		   	iCounter = iCounter + 1
	    End If
		iCount = iCount + 1

    Next

		sql = "Update tbl_Psystem_TransactionDetails Set isQuotationEntered = 1 where RequisitionId = " & PurRequisitionNo & " and ItemDescription = '" & ItemDescription & "' and ProjectId = " & ProjectId & " "
		Call DoSql(sql)

	Response.Redirect "Purchase_Quotation.asp?PurRequisitionNo=" & PurRequisitionNo

%>
<!--#include file="../includes/connection_close.asp"-->
