<!-- #include file = "include/conndb.asp" -->
<!-- #include file = "include/checkuser.asp" -->
<%
dim strSQL, rsObj, rsProduct
dim intID
intID  =  RealString(Request.QueryString("id"))
if not IsNumeric(intID) then
	Response.Write "<script language = Javascript>"
	Response.Write "alert('input again£¡');"
	Response.Write "window.history.go(-1);"
	Response.Write "</script>"
	Response.End 
end if	

	strSQL = "SELECT * FROM Product WHERE ProductType = "&intID
	set rsProduct = Server.CreateObject("ADODB.RecordSet")
	rsProduct.open strSQL,conn,1,1
	if rsProduct.eof then
		strSQL = "DELETE FROM ProductType WHERE id = "&intID
		conn.Execute ( strSQL )
%>
		<script language = Javascript>
		<!--
		alert("successfully deleted!");
		this.document.location = "manuList.asp";
		-->
		</script>
<%
	else
%>
		<script language = Javascript>
			<!--
			alert("delete the products first!");
			window.history.go(-1);
			-->
		</script>
<%
	end if
%>
