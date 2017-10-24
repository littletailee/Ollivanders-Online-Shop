<!-- #include file = "include/sysbase.asp" -->
<!-- #include file = "../include/config.asp" -->
<%
	dim rsObj,strSQL
	dim page, myKeyword, thisUrl, i, bShowAdd, sAddPage

	page = Request.QueryString("page")
	if (page = "" or IsEmpty(page)) then page = 1

	thisUrl = "orderList.asp?true=1"
	Session("adminOldUrl") = thisUrl&"&page = "&page
	set rsObj = Server.CreateObject("ADODB.RecordSet")
	strSQL = "SELECT * FROM OrderList WHERE state = 0 "
	strSQL = strSQL & " ORDER by id desc"
	rsObj.Open strSQL, conn, 1,2
	rsObj.pagesize = conMaxPerPage
%>
<html>
<head>
<title>order </title>
<meta http-equiv = "Content-Type" content = "text/html; charset = gb2312">
<link rel = "stylesheet" href = "include/main.css" type = "text/css">
</head>

<body  text = "#000000" leftmargin = "0" topmargin = "0">
<div align = "center"><br>
</div>
<table width = "99%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "4" align = "center">
  <tr bgcolor = "#4296E7"> 
    <form method = "post" action = "memberList.asp" name = "form1">
      <td height = "28" colspan = "7" bgcolor="#B1C3D9"> 
      <div align = "center"><font color = "#FFFFFF">order management</font></div>      </td>
    </form>
  </tr>
  <tr bgcolor = "#5EA5E6"> 
    <td width = "6%" nowrap bgcolor="#EBEADB"> 
    <div align = "center"> ID</div>    </td>
    <td width = "7%" nowrap bgcolor="#EBEADB"> 
    <div align = "center">customer name</div>    </td>
    <td width = "50%" nowrap bgcolor="#EBEADB"> 
    <div align = "center">address</div>    </td>
    <td width = "12%" nowrap bgcolor="#EBEADB"> 
    <div align = "center">phone</div>    </td>
    <td width = "12%" nowrap bgcolor="#EBEADB"> 
    <div align = "center">Email</div>    </td>
    <td width = "5%" nowrap bgcolor="#EBEADB"> 
    <div align = "center">treat</div>    </td>
  </tr>
  <%
		dim rsID
		i = 1
		if not (rsObj.eof or err) then rsObj.move (page-1)*conMaxPerPage
		do while not (rsObj.eof or err) 
		rsID = rsObj("id")
%>
  <tr> 
    <td width = "6%"> 
      <div align = "center"><%=rsObj("id")%></div>    </td>
    <td width = "7%" nowrap><%=rsObj("customerName")%></td>
    <td width = "50%"><%=rsObj("address")%></td>
    <td width = "12%" nowrap><%=rsObj("phone")%></td>
    <td width = "12%" nowrap><%=rsObj("email")%></td>
    <td width = "5%" title = "showdetail" style = "cursor:hand" onClick = "Javascript:window.location = 'orderProcess.asp?id=<%=rsObj("id")%>'"  nowrap>treat</td>
  </tr>
  <%
		i = i+1
		if i>conMaxPerPage then exit do
		rsObj.MoveNext
		loop
%>
  
</table>
</body>
</html>
