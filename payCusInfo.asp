
<SCRIPT language = javascript>
function form1Onsubmit()
{
	if (document.form1.customerName.value.length == 0)
	{
		alert("Please input receiver's name");
		document.form1.customerName.focus();
		return false;
	}
	if (document.form1.address.value.length == 0)
	{
		alert("Please input address");
		document.form1.address.focus();
		return false;
	}
	if (document.form1.phone.value.length == 0)
	{
		alert("Please input phone number.");
		document.form1.phone.focus();
		return false;
	}
	}
</SCRIPT>
<%
dim strSQL, rsObj
dim customerName,address,zipcode,phone,email

if Session("memberID")<>"" then
	strSQL = "SELECT * FROM member WHERE memberID = '"&Session("memberID")&"'"
	set rsObj = conn.execute (strSQL)
	if not (rsObj.eof or err) then
		customerName = rsObj("name")
		address = rsObj("address")
		zipcode = rsObj("Zipcode")
		phone = rsObj("phone")
		email = rsObj("email")
	end if
end if
%>	
<table cellpadding = 0 cellspacing = 0 border = 0 width = 100%>
<TR>
	<TD height = "22"><b>Order info</b></TD></TR>
<TR>

<tr><td>
	
	<TABLE cellSpacing = 0 cellPadding = 0 	width = "100%"  border = 1 ID = "Table1">
	<form name = form1 onsubmit = "return form1Onsubmit()" action =payStep2.asp method = post ID = "Form1">
		<TR bgcolor = "#FFDFC6"> 
		<TD height = "24"  colSpan = 2 bgcolor = "#FFCCCC"> 
			<p align = "center">Receiver's Info </p>
		</TD>
		</TR>
		<TR> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Reicever's Name£º</p>
		</TD>
		<TD  width = 431> 
			<p> 
			<INPUT maxLength = 16 size = 13 value = "<%=customerName%>"
				name = customerName style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(255,255,255)" ID = "Text1">
			</p>
		</TD>
		</TR>
		<tr> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Address<SPAN 
				>£º</SPAN></p>
		</TD>
		<TD  width = 431> 
			<p> 
			<INPUT maxLength = 16 size = 50 value = "<%=address%>"
				name = address style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(255,255,255)" ID = "Text2">
			</p>
		</TD>
		</tr>
		<tr> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Zip Code £º</p>
		</TD>
		<TD  width = 431> 
			<p> 
			<INPUT maxLength = 16 size = 13 value = "<%=zipcode%>"
				name = code style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(255,255,255)" ID = "Text3">
			</p>
		</TD>
		</tr>
		<tr> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Phone£º</p>
		</TD>
		<TD  width = 431> 
			<p> 
			<INPUT maxLength = 16 size = 23 value = "<%=phone%>"
				name = phone style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(255,255,255)" ID = "Text4">
			</p>
		</TD>
		</tr>
		<tr> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Email£º</p>
		</TD>
		<TD  width = 431> 
			<p> 
			<INPUT maxLength = 16 size = 32 value = "<%=email%>"
				name = email style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(255,255,255)" ID = "Text5">
			</p>
		</TD>
		</tr>
		<tr> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Payment Method £º</p>
		</TD>
		<TD  width = 431> 
			<p> 
			<SELECT name = "payment" style = "background-color: rgb(255,255,255)" ID = "Select1">
				<option>cash on delivery</option>                       
				<option>alipay</option>                       
				<option>credit card</option> 
			</SELECT>
			</p>
		</TD>
		</tr>
		<tr> 
		<TD  width = 87 align = "right" nowrap> 
			<p align = "right">Remark£º</p>
		</TD>
		<TD  width = 431> 
			<p> 
			<textarea cols = "50" name = "remark" style = "border:1px solid #7D85A2; font-size: 9pt; background-color:rgb(255,255,255)" rows = "3" ID = "Textarea1"></textarea>
			</p>
		</TD>
		</tr>
		
		<TR bgColor = #FFDFC6> 
		<TD  colSpan = 2 bgcolor = "#FFFFFF"> 
			<DIV align = center> 
			<p align = "center"> 
				<input type = "button" value = "Backward" name = "B4" onclick = "javascript:window.history.go(-1)" style = "border:1px solid #7D85A2; font-size: 9pt; " ID = "Button1">
				<INPUT class = main type = submit size = 3 value = "Submit" name = Submit2 style = "border:1px solid #7D85A2; font-size: 9pt; " ID = "Submit1">
			</DIV>
		</TD>
		</TR>
	</form>
	</TABLE>

</td></tr>
</table>     

      