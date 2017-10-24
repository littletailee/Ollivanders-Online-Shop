<%
Sub ShowNewProductPreview(ShowType, ProductType, KeyWord, CurPage)
%>

<TABLE cellSpacing = 0 cellPadding = 0 width = "100%" bgColor = #ffffff border = "0" style = "border-collapse: collapse" bordercolor = "#111111" ID = "Table22">
	<TBODY>
        <TD bgcolor = "<%=conBackColor%>">
			<% call ProductList(ShowType, ProductType, KeyWord, CurPage) %>
		</TD> 
		</TR>
      <TR>
         <TD align = middle bgcolor = "<%=conBackColor%>">
      </TD></TR>
	</TBODY>
</TABLE>

<%
end sub

Sub ProductList(ShowType, ProductType, KeyWord, CurPage)
	dim TotalCount, thisUrl
	dim strSrc,strTemp, strSQL, rsObj
	dim nMaxPerPage
	dim nImageHeight
	dim i, j, k
	dim szProductName	
	nMaxPerPage  =  10000
	

	if ShowType  =  "Top" then
		strSQL  =  "SELECT top 10 * FROM product WHERE Recommend >-1 ORDER by id desc"
		nMaxPerPage  =  10
	elseif ShowType  =  "All" then
		thisUrl = "product.asp?ProductType=" & ProductType & "&Keyword = " & KeyWord
		strSQL = "SELECT * FROM product WHERE id>1 "
		if not (KeyWord = "" or IsEmpty(KeyWord) ) then
			strSQL  =  strSQL & " AND name like '%"&KeyWord&"%'"
		end if
		if ProductType > 0 then
			strSQL  =  strSQL & " AND ProductType = " & ProductType
		end if
		strSQL = strSQL & " ORDER by id desc"
	end if
	
	set rsObj = conn.execute (strSQL)

	TotalCount  =  rsObj.recordCount
	
	Response.Write "<table width = '100%' border = '0' cellspacing = '0' cellpadding = '0'>"
	j = 1
	do while not (rsObj.eof or err)
		Response.Write "<td width = '100%' align = center>"			
			strSrc  =  "smallimg/"&rsObj("smallImg")
		if not (KeyWord = "" or IsEmpty(KeyWord) ) then
			szProductName  =  Replace(rsObj("name"), KeyWord, "<font color = red>" & KeyWord & "</font>")
		else
			szProductName  =  rsObj("name")
		end if		
%>
		<table  height="150" width = "98%" border = "0" cellspacing = "0" cellpadding = "0" align = "center" ID = "Table25">
			<tr> 
				<td width = 100> 
					<div align = "center"><a href = "ProductDetail.asp?id=<%=rsObj("id")%>" title = "show details">
					<img src = "<%=strSrc%>" border = 0 width="70" ></a></div>
				</td>
				<td width = "200" valign = middle><a href = "ProductDetail.asp?id=<%=rsObj("id")%>" title = "show details"><b><u><%=szProductName%></u></b></a> 
					<br>
					<strike><font color = "#666666">price: <%=rsObj("marketPrice")%> CNY</font></strike><br>
					sale price: <%=rsObj("memberPrice")%> CNY<br>
					<a href = "shopCart.asp?productID=<%=rsObj("id")%>"><img src = "images/cart.png" width = "22" height = "22" border = 0 align = "absmiddle">&nbsp;Add to cart</a>
				</td>
				<hr size='1'>
			</tr>
		</table>
<%
		Response.Write "</td></tr>"
		j = j+1
		if j>nMaxPerPage then exit do
		rsObj.MoveNext
	loop

	Response.Write "</table>"
end sub


sub ShowProduct(ProductId)
	dim strSQL, rsObj
	
	if IsEmpty(ProductId) then ProductId  =  0
	strSQL = "SELECT * FROM product WHERE id = " & ProductId
	set rsObj = conn.execute (strSQL)
	if not (rsObj.eof or err) then
%>
      <table width = "94%" border = "0"  cellspacing = "0" cellpadding = "8" align = "center" ID = "Table10">
        <tr> 
          <td width = "180" valign = "middle"> 
            <div align = "center">
              <a target = "_blank" href = "smallimg/<%=rsObj("smallImg")%>">
              <img src = "<%Response.Write "smallimg/"&rsObj("smallImg")%>" 
				width=<%=conBigImagewWidth%> border = "0" alt = "show details"></a><br>
              <br>
            </div>
          </td>
          <td valign = "middle" align="center"> <b class = "Font_10_5"><%=rsObj("name")%></b> <br>
            <br>
            <strike>price: <%=rsObj("marketPrice")%> CNY</strike><br>
            sale price: <%=rsObj("memberPrice")%>CNY<br>
            save: <font color = RED><%=rsObj("marketPrice")-rsObj("memberPrice")%></font> 
            CNY<br>
            <br>
            <a href = "shopCart.asp?productID=<%=rsObj("id")%>"><img src = "images/cart.png" width = "22" height = "22" border = "0" align = "absmiddle">&nbsp;Add to cart</a> 
          </td>
        </tr>
      </table>
       <table align = center  border = 0 cellpadding = 0 cellspacing = 0 width = "95%" ID = "Table12">
        <tbody> 
        <tr> 
          <td><hr size="1"></td>
        </tr>
        </tbody> 
      </table>
	  <table width = "94%" border = "0" cellspacing = "0" cellpadding = "4" align = "center" ID = "Table11">
        <tr> 
          <td><strong>About the product</strong></td>
		  <tr> 
          <td><p><%=convert(rsObj("introduce"))%></p></td>
          </tr>
        </tr>
      </table>
    <table align = center  border = 0 cellpadding = 0 cellspacing = 0 width = "95%" ID = "Table16">
        <tbody> 
        <tr> 
          <td><hr size="1"></td>
        </tr>
        </tbody> 
      </table>
      <table width = "94%" border = "0" cellspacing = "0" cellpadding = "4" align = "center" ID = "Table15">
        <tr> 
          <td><strong>Remark:</strong></td>
        </tr><tr> 
          <td><%=convert(rsObj("Remark"))%></td>
        </tr>
      </table>

        <table align = center  border = 0 cellpadding = 0 cellspacing = 0 width = "95%" ID = "Table16">
        <tbody> 
        <tr> 
          <td><hr size="1"></td>
        </tr>
        </tbody> 
      </table>
      <br>
      <p align = "center"><a href = "Javascript:window.history.go(-1)">Return</a> 
      </p>

<%
		dim nID
		nID  =  RealString(Request.QueryString("id"))
		strSQL = "Update product set hitNum = hitNum+1 WHERE id = " & nID
		conn.execute (strSQL)
	end if 
	set rsObj  =  nothing
	
end sub
%>

