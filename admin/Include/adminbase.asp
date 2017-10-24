

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
				'关闭记录集合和数据库连接
				rsObj.Close
				Set rsObj  =  Nothing
				CloseConn()
				Response.Clear 
				Server.Transfer "default.asp"
			end if
		end if
		'关闭记录集合和数据库连接
		rsObj.Close
		Set rsObj  =  Nothing
		CloseConn()
		
		'用户不存在或密码不正确
		Response.Write "<script language = Javascript>"
		Response.Write "alert('错误！请重新输入');"
		Response.Write "window.history.go(-1);"	
		Response.Write "</script>"
	end if
end sub

'显示管理员登录

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
