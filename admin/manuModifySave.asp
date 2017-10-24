<!-- #include file = "include/sysbase.asp" -->
<%
	   dim cmdObj, rsObj
	   dim szName, nID
	   szName  =  RealString(Request.Form("name"))
	   nID  =  RealString(Request.Form("ID"))
       Set cmdObj  =  Server.CreateObject("ADODB.Command")
       Set rsObj  =  Server.CreateObject("ADODB.RecordSet")
       cmdObj.CommandText  =  "SELECT * FROM ProductType WHERE id = " & nID
       cmdObj.CommandType  =  1
       Set cmdObj.ActiveConnection  =  conn
       rsObj.Open cmdObj, , 2,3
       rsObj("name")  =  szName
       rsObj.Update

       rsObj.Close
       set rsObj = nothing
       
       CloseConn()
%>
<script language = Javascript>
<!--
	alert("modified£¡");
	this.document.location = "manuList.asp?id=<%=nID%>";
-->
</script>
