<!-- #include file = "include/sysbase.asp" -->
<%
	   dim cmbObj, rsObj
       Set cmbObj  =  Server.CreateObject("ADODB.Command")
       Set rsObj  =  Server.CreateObject("ADODB.RecordSet")
       cmbObj.CommandText  =  "SELECT * FROM ProductType WHERE (id is null)"
       cmbObj.CommandType  =  1
       Set cmbObj.ActiveConnection  =  conn
       rsObj.Open cmbObj, , 2,3
       rsObj.AddNew 
       rsObj("name")  =  Request.Form("name")
       rsObj.Update

       rsObj.Close
       set rsObj = nothing
       CloseConn()
%>
<script language = Javascript>
<!--
	alert("successfully added!");
	this.document.location = "manuList.asp";
-->
</script>
