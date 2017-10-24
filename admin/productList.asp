<!-- #include file = "include/sysbase.asp" -->
<!-- #include file = "include/config.asp" -->
<%

	function TypeToSelect(nTypeId, szCtlName)
		dim strSQL
		dim rsType
		dim tempi
		dim nID, szName
		Response.Write	"<SELECT name =""" & szCtlName & """>"
		Response.Write	"<option value = ""0"">all categories</option>"
		
		strSQL = "SELECT * FROM ProductType "
		set rsType = Server.CreateObject("ADODB.RecordSet")
		rsType.Open strSQL, conn,1,2 
		if rsType.EOF And rsType.BOF then
			Response.Write	"</SELECT>"
			rsType.Close()
			Set rsType = Nothing
			exit function
		else
			do while not rsType.eof 
				nID = rsType("id")
				szName = rsType("name")
				Response.Write "<option value = "& nID
				
				if CInt(nID) = CInt(nTypeId) then Response.Write " selected"
				Response.Write ">"
				Response.Write szName &"</option>"
				rsType.MoveNext
			loop
		end if
		Response.Write	"</SELECT>"
		rsType.Close()
		Set rsType = Nothing
		
	end function


	dim ProductTypeDic, strSQL, rsObj, rsType
	dim thisUrl
	Set ProductTypeDic  =  Server.CreateObject("Scripting.Dictionary")
	strSQL = "SELECT * FROM ProductType "
	set rsType = conn.Execute (strSQL)
	dim tempID, tempName
	do while not (rsType.eof or err)
		tempID = CStr(rsType("id"))
		tempName = rsType("Name")
		ProductTypeDic.Add tempID, tempName
		rsType.MoveNext
	loop
	rsType.Close()
	Set rsType = Nothing
	

	dim page, myKeyword, ProductType
	myKeyword = RealString(Request.Form("myKeyword"))
	ProductType = RealString(Request.Form("ProductType"))
	if ProductType = "" then ProductType  =  "0"
	page = Request.QueryString("page")
	if (page = "" or IsEmpty(page)) then page = 1
	thisUrl = "productList.asp?ProductType=" & ProductType & "&myKeyword=" & myKeyword
	Session("adminOldUrl") = thisUrl&"&page = "&page

	strSQL = "SELECT * FROM product WHERE 1 = 1"
	if not (myKeyword = "" or IsEmpty(myKeyword) ) then
		strSQL = strSQL & " AND name like '%" & myKeyword& "%' "
	end if
	if ProductType>0 then
		strSQL = strSQL & " AND ProductType =" & ProductType
	end if
	strSQL = strSQL & " ORDER by id desc"

	set rsObj = Server.CreateObject("ADODB.RecordSet")
	rsObj.Open strSQL, conn, 1,2
	rsObj.pagesize = conMaxPerPage
%>

<html>
<head>
	<title>management</title>
	<meta http-equiv = "Content-Type" content = "text/html; charset = gb2312">
	<link rel = "stylesheet" href = "include/main.css" type = "text/css">
	<script language = "JavaScript">
		<!--
		
		function deleteMe(i){
			if(confirm("sure to delete?") == 1){
				window.location = "proDel.asp?id=" + i;
			}
		}
		
		// -->
	</script>
	<script language = Javascript src = "include/opennew.js"></script>
</head>

<body  text = "#000000" leftmargin = "0" topmargin = "2">
<div align = center>product list</div>
<table width = "98%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "4" align = "center">
  <tr> 

  </tr>
  <tr bgcolor = "#5EA5E6"> 
    <td width = "6%" nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">ID</font></div>    </td>
    <td width="14%"  nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">Name</font></div>    </td>
    <td width = "11%" nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">category</font></div>    </td>
    <td width = "10%" nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">sales price(гд)</font></div>    </td>
    <td width = "9%" nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">price(гд)</font></div>    </td>
    <td width = "7%" nowrap bgcolor="#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">buynumber</font></div>    </td>
	<td width = "15%" nowrap bgcolor = "#B1C3D9"> 
    <div align = "center"><font color = "#FFFFFF">delete</font></div>    </td>
  </tr>
<%
	dim nID, i
	dim bShowAdd, sAddPage
	i = 1
	if not (rsObj.eof or err) then rsObj.move (page-1)*conMaxPerPage
	do while not (rsObj.eof or err) 
		nID = rsObj("id")
%>
    <tr> 
		<td width = "6%"><%=nID%>&nbsp;</td>
		<td  style = "cursor:hand" title = "show details" onClick = "Javascript:window.location = 'proModify.asp?id=<%=rsObj("id")%>'" ><%=rsObj("name")%>&nbsp;</td>

		<td width = "13%"> 
<%
		Response.Write ProductTypeDic.item(cstr(rsObj("ProductType")))	
%>	  </td>
		<td width = "9%"><%=rsObj("memberPrice")%>&nbsp;</td>
		<td width = "7%"><%=rsObj("marketPrice")%>&nbsp;</td>
		<td width = "6%"><%=rsObj("buyNum")%></td>
	<td width="6%"  align=center  nowrap > 
      <a href="javascript:deleteMe(<%=nID%>)">delete</a>    </td>
  </tr>
<%
		i = i+1
		if i>conMaxPerPage then exit do
		rsObj.MoveNext
	loop
%>
  <tr bgcolor = "#4296E7"> 
    <td colspan = "9" bgcolor="#B4C5DA" height="50"> 
    
    	<% bShowAdd  =  True 
		   sAddPage  =  "proAdd.asp"
		%>
    	<%
			rsObj.Close()
			Set rsObj  =  Nothing
			CloseConn()
		%></td>
  </tr>
</table>
</body>
</html>
