 <%

dim orderID	
Sub SubmitOrder()
	dim cmdObj, rsObj, strSQL
	if Session("productList") = "" then 
		Response.Clear()
		Server.Transfer "default.asp"
	end if
		
	Set cmdObj  =  Server.CreateObject("ADODB.Command")
	Set rsObj  =  Server.CreateObject("ADODB.RecordSet")
	cmdObj.CommandText  =  "SELECT top 1 * FROM OrderList ORDER by id desc"
	cmdObj.CommandType  =  1
	
	Err.Clear()
	On Error Resume Next
	conn.BeginTrans
	Set cmdObj.ActiveConnection  =  conn
	rsObj.Open cmdObj, , 2, 3
	rsObj.AddNew 
	if Session("memberID")<>"" then
		rsObj("memberID")  = Session("memberID")
	else
		rsObj("memberID")  = 0 
	end if
	rsObj("customerName")  =  Request.Form("customerName")
	rsObj("address")  =  Request.Form("address")
	rsObj("Zipcode")  =   Request.Form("Zipcode")
	rsObj("phone")  =   Request.Form("phone")
	rsObj("email")  =  Request.Form("email")
	rsObj("payment")  =  Request.Form("payment")
	rsObj("remark")  =  Request.Form("remark")
	rsObj.Update
	rsObj.Close
	strSQL = "SELECT top 1 * FROM OrderList ORDER by id desc"
	'Response.Write  "aa<br>"
	set rsObj = conn.Execute (strSQL)
	if rsObj.eof or Err then
		Response.Write "bb<br>"
		Response.Write Err.Description 
		conn.RollbackTrans()
		Response.Write "wrong operation<a href = 'Javascript:window.history.go(-1)'>return</a>"
		Response.End
	else
		orderID = rsObj("id")
		'Response.Write "cc<br>"
		'Response.Write orderID
	end if
	rsObj.Close()
	
	dim Sum, productList, quantityArray, productArray, quantity
	dim i

	productList = Session("productList")
	If Len(productList) <>0 Then
		quantityArray  =  Split(Session("quantityList"), ", ")
		productArray  =  Split(Session("productList"), ", ")
		for i = 0 to UBound(productArray)
			strSQL  =  "SELECT * FROM product WHERE id = "&productArray(i)
			rsObj.Open strSQL, conn, 2, 1 
 			if Not rsObj.EOF or err then
				quantity  =  quantityArray(i)
				If quantity <=  0 Then quantity  =  1
				strSQL  =  "INSERT INTO orderDetail (orderID, productID, productName, price,quantity) "  
				strSQL  =  strSQL & "VALUES( "& orderID & ","
				strSQL  =  strSQL & productArray(i) &",'"
				strSQL  =  strSQL & rsObj("name")& "',"
				strSQL  =  strSQL & rsObj("memberPrice")& ","
				strSQL  =  strSQL & quantity&")"
				Conn.Execute(strSQL)
				strSQL = "Update product set buyNum = buyNum+1 WHERE id = "&productArray(i)
				conn.execute (strSQL)
			end if
		next
	else
		Err.Raise  
	end if
	if Err then 
		conn.RollbackTrans()
		Response.Write "something wrong<br>"
		Response.Write Err.Description 
	else
		conn.CommitTrans()
		Session("productList")  =  ""
		Session("quantityList")  =  ""
	End if
	rsObj.Close()
	Set rsObj = Nothing
	Set cmdObj = Nothing
	CloseConn()
	if Err then
		Response.End
	end if
	
end sub

 SubmitOrder()
%>

      <div align = "center">
        <p><br>
          <br>
          <br>
          <b><font size = "4" color = "#FF0000">Order has been submitted, please remember </font></b></p>
        <p><b><font size = "4" color = "#FF0000">your order ID £º<%=orderID%></font></b></p>
        <p><b><font size = "4" color = "#FF0000">for future referring. </font></b><br>
          <br>
          
          <br>
        </p>
      </div>
      <TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) cellPadding = 0 
                  width = "100%" borderColorLight = #aaaaaa border = 1 ID = "Table1">
        <form name = form1 action = payStep2.asp method = post ID = "Form1">               
          <tr bgcolor = "#FFDBBF"> 
            <td bgcolor = "rgb(210,232,255)"> 
              <div align = center> 
                <p align = "center"> 
                  <input type = button value = "Continue Shopping" name = Submit22 onClick = "window.location = './';" style = "border: 1px solid #7D85A2; background-color: rgb(210,232,255)" ID = "Button1">
              </div>
            </td>
          </tr>
        </form>
      </table>