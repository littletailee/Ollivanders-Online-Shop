

<% 


Sub CheckAdminLogin(strUserId, strPwd)
	dim strSQL, rsObj
	
	if strUserId <> "" then
		strSQL = "SELECT * FROM Admins WHERE Account = '" & strUserId & "'"
		
		set rsObj  =  conn.Execute(strSQL)
		if not (rsObj.eof or err) then
			if strPwd  =  rsObj("Pwd") then
				Session("AdminAccount") = rsObj("Account")
				Session("userClass") = 1
				'�رռ�¼���Ϻ����ݿ�����
				rsObj.Close
				Set rsObj  =  Nothing
				CloseConn()
				Response.Clear 
				Server.Transfer "default.asp"
			end if
		end if
		'�رռ�¼���Ϻ����ݿ�����
		rsObj.Close
		Set rsObj  =  Nothing
		CloseConn()
		
		'�û������ڻ����벻��ȷ
		Response.Write "<script language = Javascript>"
		Response.Write "alert('��������������');"
		Response.Write "window.history.go(-1);"	
		Response.Write "</script>"
	end if
end sub

'��ʾ����Ա��¼

Sub ShowAdminLogin()
%>
<form method = "post" action = "login.asp" name = "form1" ID = "Form1">
  <table width = "40%" border = "1" bordercolordark = #9CC7EF bordercolorlight = #145AA0 cellspacing = "0" cellpadding = "3" align = "center" ID = "Table4">
    <tr> 
    <td colspan = "2" height = "28" bgcolor = "#FFCCCC"> 
      <div align = "center"><font color = "#FFFFFF">login</font></div>
    </td>
  </tr>
  <tr> 
      <td width = "34%"> 
        <div align = "right">ID</div>
    </td>
      <td width = "66%"> 
        <input type = "text" name = "userID" size = "20" ID = "Text1">
    </td>
  </tr>
  <tr> 
      <td width = "34%"> 
        <div align = "right">PSW</div>
    </td>
      <td width = "66%"> 
        <input type = "password" name = "password" size = "20" value = "" ID = "Password1">
    </td>
  </tr>
  <tr> 
    <td colspan = "2" height = "28" > 
      <div align = "center"> 
        <input type = "submit" name = "Submit" value = "login" ID = "Submit1">
      </div>
    </td>
  </tr>
</table>
</form>

<%
end Sub
%>
