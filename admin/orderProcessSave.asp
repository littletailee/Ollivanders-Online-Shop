<!-- #include file = "include/sysbase.asp" -->
<%
	dim strSQL, cmdObj, rsObj
	dim nID
	nID = RealString(Request.Form("id"))
	Set cmdObj = Server.CreateObject("ADODB.Command")
	Set rsObj = Server.CreateObject("ADODB.RecordSet")
	cmdObj.CommandText = "SELECT * FROM OrderList WHERE id = "& nID
	cmdObj.CommandType = 1
	Set cmdObj.ActiveConnection = conn
	rsObj.Open cmdObj, , 2,3
	rsObj("state") = 1
	rsObj("treatedDate") =now()
	rsObj("treatedRemark") = Request.Form("treatedRemark")	
	rsObj.Update
	rsObj.Close
	set rsObj = nothing
	set cmdObj = nothing
	CloseConn()
%>
<script language = javascript>
	<!--
	alert("successfully dealt£¡");
	this.document.location = "<%=Session("adminOldUrl")%>";
	-->
</script>
