<!-- #include file = "include/conndb.asp" -->
<%
dim szMemberID
dim strSQL, rsObj, strPwd
szMemberID  =  RealString(Request.Form("memberID"))
strPwd  =  RealString(Request.Form("password"))

if szMemberID = "" then
%>
	<script language = Javascript>
	alert("ID can't be empty£¡");
	window.history.go(-1);	
	</script>
<%
	response.End
end if
strSQL = "SELECT * FROM Member WHERE MemberID = '" & szMemberID & "'"
set rsObj  =  conn.Execute(strSQL)
		if not (rsObj.eof or err) then
			if strPwd  =  rsObj("Pwd") then
				session("memberID") = rsObj("memberID")
				response.redirect "default.asp"
			else
				Response.Write "<script language = Javascript>"
				Response.Write "alert('wrong PSW');"
				Response.Write "window.history.go(-1);"	
				Response.Write "</script>"
				
			end if
else
%>
	<script language = Javascript>
	alert("ID or PSW is wrong, input again£¡");
	window.history.go(-1);	
	</script>
<%	
	end if
%>
