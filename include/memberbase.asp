
<%
function ShowMemberLogin2()
%>
	<TABLE cellSpacing = "0" cellPadding = "0" width = "100%" border = "0" ID = "Table8">
		<TBODY>
		<!--
			<TR>
				<TD height = "15" bgcolor = "<%=conGlobalBgColor%>"></TD>
			</TR>
		-->
			<TR>
				<TD height = "20" background = "images/titlebg.gif" align = center><strong>Login</strong>
				</TD>
			</TR>
			<TR>
				<TD height = "5"></TD>
			</TR>
			<TR>
				<TD vAlign = "top">
					<table width = "95%" border = "0" align = "center" cellpadding = "0" cellspacing = "0" ID = "Table9">
						<tr>
							<td>
								<% call ShowUserLogin() %>
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
<% end function 
%>		

<%


sub ShowUserLogin()
	dim strLogin
	If Session("memberID") = "" Then
	%>
	
	<table width = '100%' border = '0' cellspacing = '0' cellpadding = '0' ID = "Table3">
		<form action = "login.asp" method = "post" name = "userLogin" ID = "Form3">
			<TBODY>
				<tr>
					<td height = '25' align = 'right'>ID£º</td>
					<td height = '25'><input type = "text" name = "memberID" size = "10" maxlength = "16" ID = "Text2" class = "editbox1"></td>
				</tr>
				<tr>
					<td height = '25' align = 'right'>PSW£º</td>
					<td height = '25'><input type = "password" name = "password" size = "10" maxlength = "16" ID = "Password2" class = "editbox1"></td>
				</tr>
				<tr align = 'center'>
					<td height = '25' colspan = '2'>
						<input name = "Submit" type = "submit" value = "Login"  ID = "Submit2" >
						<input name = 'Reset' type = 'reset' id = 'Reset' value = ' Reset '>
					</td>
				</tr>
				<tr>
					<td height = '20' align = 'center' colspan = '2'><a href = 'regStep1.asp'>Enroll</a></td>
				</tr>
		</TBODY>
		</form>
	</table>
	<%
	Else %> 
		Welcome£¡<%=Session("memberID")%> <br><br>
		<a href = 'logout.asp' >logout</a>
	<%		
	end if
end sub

sub ShowMemberLogin()
	dim strLogin
	If Session("memberID") = "" Then
	%>
	
	<table width = '100%' border = '0' cellspacing = '0' cellpadding = '0' ID = "Table1">
		<form action = "login.asp" method = "post" name = "userLogin" ID = "Form1">
			<TBODY>
			<tr>
				<td>
					ID£º<input type = "text" name = "memberID" size = "10" maxlength = "16" ID = "Text1" class = "editbox1">
					PSW£º<input type = "password" name = "password" size = "10" maxlength = "16" ID = "Password1" class = "editbox1">
					<input name = "Submit" type = "submit" value = "login"  ID = "Submit1" class = "editbox1">
					<input name = 'Reset' type = 'reset' id = "Reset1" value = 'Reset' class = "editbox1">&nbsp;
					<a href = 'regStep1.asp'>Enroll</a>&nbsp;
				</td>
			</tr>
		</TBODY>
		</form>
	</table>
	<%
	Else %> 
		Welcome£¬<font color = red><b><%=Session("memberID")%></b></font>£¡ &nbsp;&nbsp;
		<a href = 'logout.asp' >Logout</a>
	<%		
	end if
end sub
%>
