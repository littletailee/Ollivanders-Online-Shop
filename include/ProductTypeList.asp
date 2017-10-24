<style type="text/css">
<!--
.STYLE5 {
	font-family: Arial, Helvetica, sans-serif;
	font-weight: bold;
}
-->
</style>
<TABLE cellSpacing = "0" cellPadding = "0" width = "100%" border = "0" ID = "Table8">
	<TBODY>
		<TR>
			<TD valign="middle" height = "30"  align = "center" bgcolor="#FFCCCC" class="STYLE5"><h2>Category</h2></TD>
		</TR>
		<TR>
			<TD height = "5"></TD>
		</TR>
		<TR>
			<TD vAlign = "middle">
				<table width = "95%" border = "0" align = "center" cellpadding = "0" cellspacing = "0" ID = "Table9">
					<tr>
					  <td>
							<% call category() %>
						</td>
					</tr>
				</table>
			</TD>
		</TR>
		<TR>
			<TD height = "5"></TD>
		</TR>
	</TBODY>
</TABLE>

<%
Sub category()
	dim strSQL, strTemp, rsObj, i
	
	Response.Write "<table width = '100%' border = '0' cellspacing = '0' cellpadding = '0' ID = 'Table3'>"
	strSQL = "SELECT * FROM ProductType "
	i = 0
	set rsObj = conn.Execute(strSQL)
	do while not (rsObj.eof or err)
		Response.Write "<tr><td height = '50' width = '500' align = 'right'>"
		Response.Write "<br><a href = Product.asp?producttype=" & rsObj("ID") & ">  <img height='70' width='90' src=images/" & rsObj("image") & ">&nbsp;&nbsp;</a></br>"
		Response.write "</td><td>"
		Response.Write "<h2><a href = Product.asp?producttype=" & rsObj("ID") & ">  " & rsObj("Name") & "</a></h2>"
		Response.write "</td></tr>"
		i = i+1
		if i>100 then exit do
		rsObj.MoveNext
	loop 
	Set rsObj = Nothing
	Response.Write "</table>"
end sub
%>