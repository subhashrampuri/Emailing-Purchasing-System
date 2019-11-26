
<!--#include file="../includes/main_page_header.asp"-->
<!--#include file="../includes/main_page_open.asp"-->


<%
 SupplierName = Replace(Server.HTMLEncode(Request.form("hdSupName")),"'","''")
 'Response.write SupplierName
 sql = "Select SupplierName from tbl_Psystem_Supplier where SupplierName = '"& SupplierName &"' "
 call RunSQL(sql,rsSupplier)

 if rsSupplier.Eof = false then
 	Sup="Y"
  else
 	Sup="N"
  end if


%>
<script language="javascript">
	function redirect()
	{
		document.Supplier.method="Post";
		document.Supplier.action="Suppliers.asp"
		document.Supplier.submit();
	}
</script>
<html>
<body onLoad="javascript:redirect();">
<form name="Supplier">
 <input type="hidden" name="hdSupplier" value="<%=SupplierName%>">
 <input type ="hidden" name="hdSup" value="<%=Sup%>">
</form>
</body>
</html>
<!--#include file="../includes/main_page_close.asp"-->