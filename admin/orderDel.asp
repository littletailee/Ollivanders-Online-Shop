<!-- #include file = "include/sysbase.asp" -->
<%
dim intId
dim strSQL
intId = RealString(Request.QueryString("id"))
strSQL = "DELETE FROM OrderList WHERE id = "&intId
'Response.Write strSQL
'Response.End
conn.execute (strSQL)
%>
<script language = Javascript>
<!--
alert("deleted£¡");
window.location = "<%=Session("adminOldUrl")%>"
//-->
</script>