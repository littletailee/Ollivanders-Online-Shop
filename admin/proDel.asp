<!-- #include file = "include/sysbase.asp" -->
<%
	dim nID
	dim strSQL
	if not IsNumeric(nID) then
		Response.Write "<script language = Javascript>"
		Response.Write "alert('wrong input!');"
		Response.Write "window.history.go(-1);"
		Response.Write "</script>"
		Response.End 
	end if	
	nID = RealString(Request.QueryString("id"))
	strSQL = "DELETE FROM Product WHERE id = " & nID 
	conn.Execute ( strSQL )
%>
<script language = Javascript>
<!--
alert("deleted");
this.document.location = "<%=Session("adminOldUrl")%>";
-->
</script>