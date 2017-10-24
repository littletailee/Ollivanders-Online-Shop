<!-- #include file = "include/sysbase.asp" -->
<%
  Dim strSQL, cmdObj, rsObj
  Dim nID, nProductType, nMemberPrice, nMarketPrice
  nID  =  RealString(Request.Form("id"))
  nProductType  =  Cint(Request.Form("ProductType"))
  nMemberPrice  =  Request.Form("memberPrice")
  nMarketPrice  =  Request.Form("marketPrice")
  if not IsNumeric(nMemberPrice) then nMemberPrice  =  "1"
  if not IsNumeric(nMarketPrice) then nMarketPrice  =  "1"
  Set cmdObj  =  Server.CreateObject("ADODB.Command")
  Set rsObj  =  Server.CreateObject("ADODB.RecordSet")
  cmdObj.CommandType  =  1
  cmdObj.CommandText  =  "SELECT top 1 * FROM product WHERE id = " & nID
  cmdObj.ActiveConnection  =  conn
  call rsObj.Open(cmdObj, , 2,3)
  
  rsObj("ProductType")  =  Request.Form("ProductType")
  rsObj("name")  =  Request.Form("name")
  rsObj("memberPrice")  =  nMemberPrice
  rsObj("marketPrice")  =  nMarketPrice
  rsObj("introduce")  =  Request.Form("introduce")
  rsObj("remark")  =  Request.Form("remark")
  rsObj.Update()
        rsObj.Close
	   
       set rsObj = nothing
       CloseConn()
%>
<script language = Javascript>
	<!--
	alert("modified!");
	window.location = "<%=Session("adminOldUrl")%>"
	-->
</script>
