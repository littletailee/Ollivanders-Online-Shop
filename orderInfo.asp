<%
dim rsObj,strSQL
dim orderID

orderID = RealString(Request.Form("orderID"))
set rsObj = Server.CreateObject("ADODB.RecordSet")
if orderID = "" then orderID = 0

strSQL = "SELECT * FROM OrderList WHERE id = "& orderID &" AND memberID = '"&Session("memberID")&"'"
rsObj.Open strSQL, conn, 2,1

%>
<TABLE cellSpacing = 0  cellPadding = 0 
                  width = "100%" borderColorLight = #aaaaaa border = 1 ID = "Table1">
        <tr bgcolor = "#FFDFC6"> 
          <td colspan = "2" nowrap bgcolor = "#FFCCCC"> 
            <div align = "center">Order ID:<%=orderID%> </div>
          </td>
        </tr>
<%if not (rsObj.eof or err) then %>
        <tr>        
        <tr> 
          <td width = "12%" height="30" nowrap > 
          <div align="left">order date</div></td>
          <td width = "88%"><%=rsObj("createDate")%>¡¡</td>
        </tr>
        <tr> 
          <td width = "12%" height="27" nowrap > 
          <div align="left">delivery date</div></td>
          <td width = "88%"><%=rsObj("treatedDate")%>¡¡</td>
        </tr>
        <tr> 
          <td width = "12%" height="24" nowrap > 
          <div align="left">remark</div></td>
          <td width = "88%"><%=rsObj("treatedRemark")%>¡¡</td>
        </tr>
</table>
<%
strSQL = "SELECT * FROM orderDetail WHERE orderID = "& orderID
set rsObj = conn.execute (strSQL)
%>
      <TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) cellPadding = 0 
                  width = "100%" borderColorLight = #aaaaaa border = 1 ID = "Table2">
        <tr bgcolor = "#FFEFB0"> 
          <td width = "10%" bgcolor = "#FFCCCC"> 
            <div align = "center">ID</div>
          </td>
          <td width = "46%" bgcolor = "#FFCCCC"> 
            <div align = "center">Product Name</div>
          </td>
          <td width = "13%" bgcolor = "#FFCCCC"> 
            <div align = "center">Price</div>
          </td>
          <td width = "14%" bgcolor = "#FFCCCC"> 
            <div align = "center">Number</div>
          </td>
          <td width = "17%" bgcolor = "#FFCCCC"> 
            <div align = "center">total</div>
          </td>
        </tr>
        <%
dim totalPrice
totalPrice = 0
do while (not rsObj.eof or err)  
%>
        <tr> 
          <td width = "10%"> 
            <div align = "center"><%=rsObj("productID")%>&nbsp; </div>
          </td>
          <td width = "46%"> 
            <div align = "center"><%=rsObj("productName")%>&nbsp; </div>
          </td>
          <td width = "13%"> 
            <div align = "center"><%=rsObj("price")%>&nbsp; </div>
          </td>
          <td width = "14%"> 
            <div align = "center"><%=rsObj("quantity")%>&nbsp; </div>
          </td>
          <td width = "17%"> 
            <div align = "center">
              <%if rsObj("price")<>"" AND rsObj("quantity")<>"" then Response.Write rsObj("price")*rsObj("quantity")%>
              &nbsp; </div>
          </td>
        </tr>
        <%
  	if rsObj("price")<>"" AND rsObj("quantity")<>"" then
	  	totalPrice = totalPrice+rsObj("price")*rsObj("quantity")
	end if
	rsObj.MoveNext
loop
%>
        <tr bgcolor = "#FFEFB0"> 
          <td colspan = "5" bgcolor = "#FFFFFF"> 
            <div align = "right">Total£º<%=totalPrice%>CNY&nbsp;&nbsp;&nbsp;</div>
          </td>
        </tr>
        <tr bgcolor = "#FFCC00"> 
          <td colspan = "5" bgcolor = "#FFFFFF"> 
            <div align = "center">
              <input type = "button" name = "Submit3" value = "Backward" onClick = "window.history.go(-1);" style = "border: 1px solid #7D85A2; " ID = "Button1">
            </div>
          </td>
        </tr>
</table>
<%
else 
	Response.Write "</table><br><center>no results, please search again after login</center>"
end if %></TD></TR>
                                </TBODY></TABLE>