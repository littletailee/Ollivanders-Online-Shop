<!-- #include file = "include/sysbase.asp" -->
<!-- #include file = "../include/config.asp" -->
<%
    Dim rsObj, rsSuper, strSQL
    Dim page, thisUrl, i, bShowAdd, sAddPage
    Dim intID, nSuperId, szSuperName, nID, szName
    intID  =  RealString(Request.QueryString("id"))
    If intID  =  "" Then intID  =  "0"
    page  =  RealString(Request.QueryString("page"))
    If (page  =  "" Or IsEmpty(page)) Then page  =  1
	thisUrl  =  "manuList.asp?id=" & intID
	Session("adminOldUrl") = thisUrl&"&page = "&page
	Set rsObj  =  Server.CreateObject("ADODB.RecordSet")
	strSQL  =  "SELECT * FROM ProductType WHERE 1 = 1  ORDER by id"	
	rsObj.Open strSQL, conn, 1,2
	rsObj.pagesize  =  conMaxPerPage
	strSQL  =  "SELECT * FROM ProductType WHERE id = " & intID
	Set rsSuper  =  Server.CreateObject("ADODB.RecordSet")
	rsSuper.Open strSQL, conn, 1,2
	If Not rsSuper.eof Then
		szSuperName  =  rsSuper("Name")
    End If
    rsSuper.Close()
%>
<html><head>
	<meta http-equiv = "Content-Type" content = "text/html; charset = gb2312">
	<link rel = "stylesheet" href = "include/main.css" type = "text/css">
	<SCRIPT language = "JavaScript" src = "Include/common.js"></SCRIPT>
</head>
<body  text = "#000000" leftmargin = "0" topmargin = "0" marginwidth = "0" marginheight = "0">
<p>
<table align = center width = "60%" border = "1" bordercolordark = "#33CC00" 
bordercolorlight = "#092094" cellpadding = "2" cellspacing = "0" align = "center">
<tr>
      <td height = "28" bgcolor = "#B8C8DC"> 
        <div align = "center"><font color = "#FFFFFF">category management</font></div>
      </td>
</tr>
</table>
<%	
	
	if intID<>"0" then 
%>
<table width = "60%" border = "1" bordercolordark = "#33CC00" bordercolorlight = "#092094" cellpadding = "2" cellspacing = "0" align = "center">
  <tr> 
      <td width = "7%" nowrap height = "37" valign = "middle"> <a href = manuList.asp?id=<%=nSuperId%>>backwards</a></td>
      <td width = "93%" height = "37" valign = "middle" align = center> 
		<form method = "post" name = "sortMod" action = "manuModifySave.asp">category
			<input type = "text" name = "name" value = "<%=szSuperName%>">
			<input type = "hidden" name = "id" value = "<%=intID%>">
			<input type = "submit" name = "Submit" value = "save" onClick = "return checkMe(sortMod)">
        </form>
      </td>
  </tr>
</table>
<%
	end if 
%> 

           
<table width = "60%" border = "1" bordercolordark = "#33CC00" bordercolorlight = "#092094" cellpadding = "2" cellspacing = "0" align = "center">
<%
if not (rsObj.eof or err) then 
%>
 <tr bgcolor = "#5EA5E6">      
    <td width = "6%" nowrap bgcolor="#B4C5DA"> 
    <div align = "center"><font color = "#FFFFFF">ID</font></div></td>
    <td  nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">category</font></div></td>
    <td width = "8%" nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">modify</font></div></td>
    <td width = "8%" nowrap bgcolor="#B8C8DC"> 
    <div align = "center"><font color = "#FFFFFF">delete</font></div></td>
 </tr>
<%
	rsObj.move (page-1)*conMaxPerPage

    i  =  1
    Do While Not (rsObj.EOF Or Err)
      nID  =  rsObj("id")
      szName  =  rsObj("Name")
%>
      <tr>
		<td width = "6%" nowrap><div align = "center"><%=nID%></div></td>
		<td  nowrap> 
		<div align = "center"><a href = manuList.asp?id=<%=nID%>><%=szName%></a></div></td>
		<td width = "6%" nowrap> 
		<div align = "center"><a href = manuList.asp?id=<%=nID%>>modify</a></div></td>
		<td width = "6%" nowrap align = center> 
		<a href = "manuDel.asp?id=<%=nID%>" onclick = "return confirm('delete?');">delete</a></td>
      </tr>
<%
		i  =  i + 1
		If i > conMaxPerPage Then Exit Do
		rsObj.MoveNext()
		loop
	end if
    bShowAdd  =  True
    sAddPage  =  "manuAdd.asp"
%>
       
</table>
<% 
	rsObj.Close()
	Set rsObj  =  Nothing
	CloseConn()
%>      
</body>
</html>
