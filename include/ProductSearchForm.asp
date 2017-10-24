<%

sub ShowProductTypeList(nTypeId)
%>
	<SELECT name = "ProductType" style = " font-size:9pt" ID = "Select1">
		<option value = "0">all categories</option>
			<%call ProductTypeToSelect(nTypeId)%>
    </SELECT>
<%
end sub


sub ProductTypeToSelect(nTypeId)
	dim strSQL, rsObj
	dim tempi
	dim nID, szName
	if Not IsNumeric(nTypeId) then
	  nTypeId = 0
	end if
	strSQL = "SELECT * FROM ProductType"
	set rsObj = conn.execute (strSQL)
	if (rsObj.eof or err) then
		Response.Write "ERr!"
		exit sub
	else
		do while not rsObj.eof 
			nID = rsObj("id")
			szName = rsObj("name")
			Response.Write "<option value = "&nID
			if CLng(nID) = CLng(nTypeId) then Response.Write " selected"
			Response.Write ">"
			Response.Write szName&"</option>"
			rsObj.MoveNext
		loop
	end if
end sub
%>

<%
sub ShowSearchForm(myKeyword, nTypeId)
%>
<table width = "100%" border = "0" cellspacing = "0" cellpadding = "4" align = "center" height = "30" ID = "Table7">
        <tr> 
          <form method = "post" name = "orderSearch" action = "orderSearch.asp" ID = "Form1">
            <td width = "45%" valign = "top"> order ID 
              <input type = "text" size = "12" name = "orderID" onKeyUp = "checknum(orderSearch.orderID)" style = "font-size: 9pt; border: 1px solid #7D85A2; " ID = "Text1">
              <input type = "submit" name = "Submit3" value = "search" class = "forInput" style = "font-size: 9pt; border: 1px solid #7D85A2; " ID = "Submit1">
            </td>
          </form>
          <form name = "productSearch" method = "post" action = "product.asp" ID = "Form2">
            <td width = "55%" valign = "top"> 
              <input type = "text" name = "Keyword" size = "12" value = "<%=myKeyword%>" style = "font-size: 9pt; border: 1px solid #7D85A2; " ID = "Text2">
              <% call ShowProductTypeList(nTypeId) %>
              <input type = "submit" name = "Submit4" value = "search" class = "forInput" style = "font-size: 9pt; border: 1px solid #7D85A2; " ID = "Submit2">
            </td>
          </form>
        </tr>
</table>
<%
end sub
%>
