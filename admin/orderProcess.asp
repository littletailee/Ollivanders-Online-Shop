<!-- #include file = "include/sysbase.asp" -->
<%
dim rsObj,strSQL
dim nOrderId
set rsObj = Server.CreateObject("ADODB.RecordSet")
nOrderId = RealString(Request.QueryString("id"))
strSQL = "SELECT * FROM OrderList WHERE id = "& nOrderId
rsObj.Open strSQL, conn, 2,1 
%>
<html>
<head>
<title>查看定单</title>
<meta http-equiv = "Content-Type" content = "text/html; charset = gb2312">
<link rel = "stylesheet" href = "include/main.css" type = "text/css">

<script language = Javascript>
<!--
function deleteMe(){
	if (confirm("delete?") == 1){
		window.location = "orderDel.asp?id=<%=RealString(Request.QueryString("id"))%>";
	}
}

function submitMe(){
	if (confirm("deliever?") == 1){
		form1.submit();
	}
}
//-->
</script>
</head>

<body  text = "#000000" leftmargin = "0" topmargin = "6">
<table width = "98%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "2" align = "center">
  <tr bgcolor="#33CC00"> 
    <td height = "28" colspan = "2" bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">order detail </font></div>    </td>
  </tr>
  <tr bgcolor = "#33CC00"> 
    <td colspan = "2" nowrap bgcolor="#EBEADB">
      <font color = "#660033"> order detail</font></td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">order ID</div>
    </td>
    <td width = "88%"><%=rsObj("id")%></td>
  </tr> 
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">receiver's name</div>
    </td>
    <td width = "88%"><%=rsObj("customerName")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">address</div>
    </td>
    <td width = "88%"><%=rsObj("address")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">zip code</div>
    </td>
    <td width = "88%"><%=rsObj("Zipcode")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">phone</div>
    </td>
    <td width = "88%"><%=rsObj("phone")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">email</div>
    </td>
    <td width = "88%"><%=rsObj("email")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">payment method</div>
    </td>
    <td width = "88%"><%=rsObj("payment")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">remark</div>
    </td>
    <td width = "88%"><%=rsObj("remark")%>&nbsp;</td>
  </tr>
  <tr> 
    <td width = "12%" nowrap> 
      <div align = "right">order date</div>
    </td>
    <td width = "88%"><%=rsObj("createDate")%>&nbsp;</td>
  </tr>
  <tr bgcolor = "#33CC00"> 
    <td colspan = "2" nowrap bgcolor="#EBEADB">
      <font color = "#660033">order details</font></td>
  </tr>
</table>
<table width = "98%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "2" align = "center">
  <tr> 
    <td width = "8%" nowrap> 
      <div align = "center">Product ID</div>
    </td>
    <td width = "49%" nowrap> 
      <div align = "center">product name</div>
    </td>
    <td width = "14%" nowrap> 
      <div align = "center">price</div>
    </td>
    <td width = "13%" nowrap> 
      <div align = "center">number</div>
    </td>
    <td width = "16%" nowrap> 
      <div align = "center">total</div>
    </td>
  </tr>
<%
dim nSinglePrice, nOrderPrice, nTotalPrice, nQuantity
nSinglePrice = 0	
nQuantity = 0		
nOrderPrice = 0		
nTotalPrice = 0		
strSQL = "SELECT * FROM orderDetail WHERE orderID = " & nOrderID
Set rsObj = conn.Execute(strSQL)
do while (not rsObj.eof or err)  
	nSinglePrice = rsObj("price")
	nQuantity = rsObj("Quantity")
  	if nSinglePrice <>"" AND nQuantity <> "" then
  		nOrderPrice = nSinglePrice * nQuantity
	  	nTotalPrice = nTotalPrice + nOrderPrice
	end if
%>
  <tr bgcolor = "#D9EAF9"> 
    <td width = "8%"> 
      <div align = "center"><%=rsObj("productID")%>&nbsp; </div>
    </td>
    <td width = "49%"> 
      <div align = "center"><%=rsObj("productName")%>&nbsp; </div>
    </td>
    <td width = "14%"> 
      <div align = "center"><%=nSinglePrice%>&nbsp; </div>
    </td>
    <td width = "13%"> 
      <div align = "center"><%=nQuantity%>&nbsp; </div>
    </td>
    <td width = "16%"> 
      <div align = "center"><%=nOrderPrice%>&nbsp; </div>
    </td>
  </tr>
<%
	rsObj.MoveNext
loop
rsObj.Close()
Set rsObj = Nothing
CloseConn()
%>
  <tr> 
    <td colspan = "5"> 
      <div align = "right">Total：<%=nTotalPrice%>CNY&nbsp;&nbsp;&nbsp;</div>
    </td>
  </tr>
</table>
  <table width = "98%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "2" align = "center">
<form method = "post" action = "orderProcessSave.asp" name = "form1">
  <tr>
      <td width = "93%"> 
        <textarea name = "treatedRemark" cols = "70" rows = "3">delievered</textarea>
        <input type = "hidden" name = "id" value = "<%=nOrderId%>">
      </td>
  </tr>
  <tr bgcolor="#33CC00"> 
    <td height = "28" colspan = "2" bgcolor="#EBEADB">
      <div align = "center">
        <input type = "button" name = "Submit3" value = "delete" onClick = "deleteMe()">
        <input type = "button" name = "Submit" value = " deliever " onClick = "submitMe()">&nbsp;
        <input type = "button" name = "Submit2" value = "backward " onClick = "window.location = '<%=Session("adminOldUrl")%>'">
      </div>    </td>
  </tr>
</form>
</table>
<br>
<br>
</body>
</html>
