<%
	str = Request.QueryString("ID")
	response.write str
%>
<html>
<head>
<title></title>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1">
<script language="javascript">
function changeOnClick(frm,iType)
{
	var myImage=Array("GRN_0.jpg","GRN_1.jpg","GRN_2.jpg");
	var imgName;
	var val;
	imgName=""
	
	val=frm.hdValue.value;

	if (iType==0)
	{
		if (frm.hdValue.value != 1) 
			{
				val=parseInt(frm.hdValue.value)-1;
				
			}
		else
			{
				val=1;
				
				
			}
	}
	else if(iType==1)
	{
		if (frm.hdValue.value !=3)
			{
				val=parseInt(frm.hdValue.value)+1;
				
			}
	
		else
			{
				val=3;
				
			}
	}
		

	imgName="Images/"+myImage[val-1];
	frm.MW.src=imgName;
	frm.hdValue.value=val;

	if (val >1 && val <=2)
	{
		frm.cmdNext.width=40;
		frm.cmdNext.height=22;
		frm.cmdPrev.width=40;
		frm.cmdPrev.height=22;
	}

	else  if (val ==1)
	{
		frm.cmdNext.width=40;
		frm.cmdNext.height=22;
		frm.cmdPrev.width=0;
		frm.cmdPrev.height=0;
	}	

	else  if (val ==3)
	{
		frm.cmdNext.width=0;
		frm.cmdNext.height=0;
		frm.cmdPrev.width=40;
		frm.cmdPrev.height=22;
	}		
		
 }

</script>
</head>
<body bgcolor="#FFFFFF" text="#000000" onLoad="changeOnClick(document.frmWiz,0)">
<form name="frmTest">
<input type="hidden" name="hdTest" value="<%=str%>">
</form>
<script language="javascript">
var val = document.frmTest.hdTest.value;
alert(val);








</script>

<form name="frmWiz">
  <table width="100%" border="0" cellspacing="1" cellpadding="1" bgcolor="#DBDBDB">
    <tr bgcolor="#FFFFFF"> 
      <td width="100%">
        <div align="center"></div>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF" align="center" valign="top"> 
      <td> 
        <div align="center"><img name="MW"></div>
      </td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td>
	  <table cellpadding="0" cellspacing="1" border="0" width="100%">
          <tr> 
            <Td width="42%"><input type="hidden" name="hdValue" VALUE="1">
            </td>
            <Td width="8%"><input type="button" name="cmdPrev" value="BACK" onClick="changeOnClick(document.frmWiz,0)"></td>
            <Td width="50%"> 
              <input type="button" name="cmdNext" value="NEXT" onClick="changeOnClick(document.frmWiz,1)"> </td>
          </tr>
        </table>
	  </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td> 
        <div align="right"> 
          <div style="margin-right:10" class="domain-text"> 
            <div align="right"><a href="#" onClick="window.close()" style="text-decoration:none"><font color="blue">Close</font></a></div>
          </div>
        </div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
