<%
	if  Session("AdminAccount")  =  "" then
		Response.Clear
		Server.Transfer ("login.asp")
	end if
%> 
