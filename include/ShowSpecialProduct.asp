
<% 
Sub ShowSpecialProduct(strFilter, nCount)
	dim strSpecialText
	if strFilter  =  "Hot" then
		strSpecialText  =  "Hot"	
	elseif strFilter  =  "Cheap" then
		strSpecialText  =  "On Sale"
	else
		strFilter  =  "Hot"
		strSpecialText  =  "Hot"
	end if
%>
<style type="text/css">
<!--
.STYLE1 {	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;}
-->
</style>

	<table cellSpacing = "0" cellPadding = "0" width = "100%" border = "0" ID = "Table1">
		<tbody>
			<tr>
				<td height = "30"  align = "center" bgcolor="#FFCCCC" class="STYLE5"><%=strSpecialText%><span class="STYLE1"></span></td>
			</tr>
			<tr>
				<td height = "5"></td>
			</tr>
			<tr>
				<td vAlign = "top">
					<table width = "95%" border = "0" align = "center" cellpadding = "0" cellspacing = "0" ID = "Table2">
						<tr>
							<td>
								<% call ShowSpecialProductList(strFilter, nCount) %>
							</td>
						</tr>
					</table>
				</td>
			</tr>
			<tr>
				<td height = "5"></td>
			</tr>
		</tbody>
	</table>
    <%
end sub

Sub ShowSpecialProductList(strFilter, nCount)
	dim strSQL, rsObj, productName, i
	nCount = CInt(nCount)
	if nCount > 20 then nCount = 20
	if nCount < 1 then nCount = 1
	
	Response.Write "<table width = '100%' border = '0' cellspacing = '0' cellpadding = '0' ID = 'Table3'>"
	if strFilter  =  "Hot" then
		strSQL = "SELECT TOP " & nCount & " ID, [Name] FROM product ORDER BY buyNum desc,id desc"
	elseif strFilter  =  "Cheap" then
		strSQL = "SELECT TOP " & nCount & " ID, [Name], 1- memberPrice/marketPrice AS Discount FROM Product WHERE marketPrice <> 0 ORDER BY memberPrice/marketPrice, id desc"
	else
		strSQL = "SELECT TOP " & nCount & " ID, [Name] FROM Product ORDER BY buyNum desc,id desc"
	end if
		
	set rsObj = conn.Execute (strSQL)
	
	if rsObj.EOF And rsObj.BOF then
		Response.Write "No recommendation¡­¡­"
	end if
	i = 1
	do while not (rsObj.eof or err)
		Response.Write "<tr><td height  =  50>"
		if strFilter  =  "Cheap" then
			productName = i & "." & rsObj("name") & "(<font color = red>" & FormatNumber(rsObj("Discount")*100,1) & "</font>%off)"
		else
			productName = rsObj("name")
			if len(productName) > 18 then
				productName  =  left(productName,16)&"..."
			end if
		end if
		Response.Write "<h2 align = 'center'><a href = ProductDetail.asp?id=" & rsObj("ID") & ">  " & productName & "</a></h2>"
		Response.Write "</td></tr>"
		i = i+1
		if i>100 then exit do
		rsObj.MoveNext
	loop 
	rsObj.Close()
	Set rsObj = Nothing
	Response.Write "</table>"
end sub
%>
