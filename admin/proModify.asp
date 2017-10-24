<!-- #include file = "include/sysbase.asp" -->
<!-- #include file = "include/ProductTypeBase.asp" -->
<%
	dim rsObj,strSQL
	dim nID, nTypeId, szName, MemberPrice, MarketPrice, szIntro, szRemark
	nID = RealString(Request.QueryString("id"))
	strSQL = "SELECT * FROM product WHERE id = " & nID
	set rsObj = conn.Execute(strSQL)
	if rsObj.eof or err then
		Response.Write "wrong,<a href = Javascript:window.history.go(-1)>return</a>"
		Response.End()
	end if
	nTypeId = rsObj("ProductType")
	szName = rsObj("Name")
	MemberPrice = rsObj("MemberPrice")
	MarketPrice = rsObj("MarketPrice")
	szIntro = rsObj("Introduce")
	szRemark = rsObj("Remark")
	rsObj.Close()
	Set rsObj = Nothing
%>
<html>
<head>
<title>modify</title>
<meta http-equiv = "Content-Type" content = "text/html; charset = gb2312">
<link rel = "stylesheet" href = "include/main.css" type = "text/css">
<script language = Javascript>
	<!--
	function checkForm(){
	if (form1.ProductType.value < 0){
		alert("choose category");
		return false;
		}
	}
	//-->
</script>
</head>

<body  text = "#000000">
<br>
<form method = "post" action = "proModifySave.asp" name = "form1" onSubmit = "return checkForm()">
  <table width = "70%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "4" align = "center">
    <tr bgcolor="#33CC00"> 
      <td colspan = "2" bgcolor="#B1C3D9"> 
      <div align = "center"><font color = "#FFFFFF">modify</font></div>      </td>
    </tr>
    <tr> 
      <td width = "20%" nowrap> 
        <div align = "right">category</div>
      </td>
      <td width = "80%">
          <%
			call TypeToSelect(nTypeId, "ProductType")
			CloseConn()
		  %>
        <input type = "hidden" name = "id" value = "<%=nID%>">
      </td>
    </tr>
    <tr> 
      <td width = "20%" nowrap> 
        <div align = "right">name</div>
      </td>
      <td width = "80%"> 
        <input type = "text" name = "name" size = "40" value = "<%=szName%>">
      </td>
    </tr>
    <tr> 
      <td width = "20%" nowrap> 
        <div align = "right">sales price</div>
      </td>
      <td width = "80%"> 
        <input type = "text" name = "memberPrice" size = "12" value = "<%=MemberPrice%>">
        гд</td>
    </tr>
    <tr> 
      <td width = "20%" nowrap> 
        <div align = "right">price</div>
      </td>
      <td width = "80%"> 
        <input type = "text" name = "marketPrice" size = "12" value = "<%=MarketPrice%>">
        гд</td>
    </tr>
    <tr> 
      <td width = "20%" nowrap> 
        <div align = "right">introduction</div>
      </td>
      <td width = "80%"> 
        <textarea name = "introduce" cols = "60" rows = "4"><%=szIntro%></textarea>
      </td>
    </tr>
    <tr> 
      <td width = "20%" nowrap> 
        <div align = "right">remark</div>
      </td>
      <td width = "80%"> 
        <textarea name = "remark" cols = "60" rows = "4"><%=szRemark%></textarea>
      </td>
    </tr>
    <tr> 
      <td colspan = "2" nowrap> 
        <div align = "center">
          <input type = "submit" name = "Submit" value = "modify">
          <input type = "button" name = "Submit2" value = "backward" onClick = "window.location = '<%=Session("adminOldUrl")%>';">
        </div>
      </td>
    </tr>
  </table>
</form>
</body>
</html>
