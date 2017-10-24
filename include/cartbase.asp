<%
Sub PutToCart( productID,quantity)
   dim productList, quantityList
   productList = Session("productList")
   quantityList = Session("quantityList")
   
   If Len(productList)  =  0 Then
      Session("productList")  = productID 
	  Session("quantityList") = quantity 
   ElseIf InStr( productList & ",",  productID & "," ) <=  0 Then
      Session("productList")  =   productID & ", " & productList
      Session("quantityList")  =  quantity  & ", " & quantityList
   End If
End Sub


sub ShowCart()
	dim canPay
%>
<style type="text/css">
<!--
.STYLE3 {font-size: large}
.STYLE8 {
	font-size: 18px;
	font-family: "Courier New", Courier, monospace;
}
-->
</style>

<table width = "100%" border = "0" cellspacing = "0" ID = "Table1">
	<TR>
		<TD height = "30"  bgColor=#FFCCFF><span class="STYLE3"><b> &nbsp;Cart List</b></b></span></TD>
	</TR>
	<TR>
  <tr>
    <td width = "80%" valign = "top"><p align = "center">
		<font class = main1 STYLE1><%=Head%></font>
		<form name = "form1" Action = "shopCart.asp" Method = "POST" ID = "Form1">
			<div align = "center">
<%	
			canPay  =  ShowCartTable()
%>
			<blockquote>
				<p align = "center">
				<input type = "button" value = "Continue Shopping" name = "B2" 
				style = "border:1px solid #7D85A2; font-size: 9pt; 
				background-color:rgb(210,232,255)" onclick = "window.location = '<%=Session("oldUrl")%>';" 
				style = "font-size: 9pt" ID = "Button1">&nbsp;&nbsp;
				<input type = "button" value = "Cancle Order" name = "B3" 
				style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(210,232,255)" 
				OnClick = "clean()" style = "font-size: 9pt" <%=canPay%> ID = "Button2">&nbsp;&nbsp;
				<input type = "button" value = "Submit Order" name = "B4" 
				style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(210,232,255)"
				 onclick = "window.location.href = 'payStep1.asp';" style = "font-size: 9pt" <%=canPay%> ID = "Button3">
			</blockquote>
			
            <br>
		</form>    </td>
  </tr>
</table>
<%  	
end sub


function ShowCartTable()
	dim Sum, canPay
	dim quantityArray, productArray, quantity, i
	dim strSQL, rsObj
%>
<table border = "1" cellpadding = "0" cellspacing = "0" width = "95%" bgcolor = "#FFFFFF" bordercolor = "#808080" style = "border-collapse: collapse">
	<tr bgcolor = "rgb(210,232,255)"> 
		<td width = "59"  height = "22" align = "center" bgcolor="#FFCCCC" >Product ID</td>
		<td width = "191"  height = "22" align = "center" bgcolor="#FFCCCC" >Product Name</td>
		<td width = "76" height = "22" align = "center" bgcolor="#FFCCCC" >Price</td>
		<td width = "77"  height = "22" align = "center" bgcolor="#FFCCCC" >Quantity</td>
		<td width = "77"  height = "22" align = "center" bgcolor="#FFCCCC"  >Total price</td>
	</tr>
<%
Sum  =  0
If Len(Session("productList")) <>0 Then
	quantityArray  =  Split(Session("quantityList"), ", ")
	productArray  =  Split(Session("productList"), ", ")
	
	for i = 0 to UBound(productArray)
		strSQL  =  "SELECT * FROM product WHERE id = "&productArray(i)
		set rsObj = conn.execute (strSQL)
		
 		if Not rsObj.EOF or err then
			quantity  =  quantityArray(i)
			If quantity <=  0 Then quantity  =  1
			Sum  =  Sum + rsObj("memberPrice") * quantity
%>
	<tr> 
		
		<td align = "center" width = "59"><%=rsObj("id")%> ¡¡</td>
		<td align = "center" width = "191"><%=rsObj("name")%> ¡¡</td>
		<td align = "center" width = "76">£¤<%=rsObj("memberPrice")%>¡¡</td>
		<td align = "center" width = "77"> 
			<input Name = "quantity<%=rsObj("id")%>" Value = "<%=quantity%>" Size = "4" onKeyUp = "checknum(quantity<%=rsObj("id")%>)" ID = "Text1">
		</td>
			<td Align = "center" width = "73">£¤<%=rsObj("memberPrice")*quantity%></td>
   </tr>
<%	
		end if
		set rsObj  =  nothing
	next
	canPay = ""
else
	canPay = "disabled"
end if
%>
	<tr> 
	  <td Align = "Right" ColSpan = "6" width = "520" height = "24"><span class="STYLE8">Total:&nbsp;<%=Sum%>&nbsp;CNY</span>	    <font color = "#ABABAB">&nbsp;&nbsp;</font></td>
    </tr>
</table>

<%
	ShowCartTable  =  canPay
end function
%>
<Script language=javascript>
	function clean(){ 
		if (confirm("Clear the cart?") == 1){
		window.location.href = "shopCart.asp?clear=yes"}
	}

	function checkNumNull(theform) {
		if (theform.value == "") {	
		alert("Please fill the number of the product");
			theform.focus();
			return false;
		}
	}
</Script>
