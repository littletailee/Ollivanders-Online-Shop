 <TABLE cellSpacing = 0 cellPadding = 0 width = "100%" bgColor = #ffffff border = "0" style = "border-collapse: collapse" bordercolor = "#111111" ID = "Table1">
                                <TBODY>
                                <TR>
                                <TD bgColor = #FFFFFF>
<%
	dim strSQL, rsObj, cmdObj
	dim strPwd1, strPwd2, szMemberId
	strPwd1  =  Request.Form("password1")
	strPwd2  =  Request.Form("password2")
	szMemberId  =  RealString(Request.Form("memberID"))
	
	if Request.Form("memberID") = "" then
%>
	<script language = Javascript>
	<!--
	alert("用户名不能为空");
	window.history.go(-1);
	//-->
	                            </script>
<%
	Response.End
end if
strSQL = "SELECT memberID FROM member  WHERE memberID = '"& szMemberId &"'"
set rsObj = conn.execute (strSQL)
if not rsObj.eof then
%>
	<script language = Javascript>
	<!--
	alert("该用户名已存在，请重新选择用户名");
	window.history.go(-1);
	//-->
	                            </script>
<%
	Response.End
end if

if strPwd1<>strPwd2 then%>
	<script language = Javascript>
	<!--
	alert("密码、确认密码不同");
	window.history.go(-1);
	//-->
	                            </script>
<%
	Response.End
end if
       Set cmdObj  =  Server.CreateObject("ADODB.Command")
       Set rsObj  =  Server.CreateObject("ADODB.RecordSet")
       cmdObj.CommandText  =  "SELECT top 1 * FROM member ORDER by memberID desc"
       cmdObj.CommandType  =  1
       Set cmdObj.ActiveConnection  =  conn
       rsObj.Open cmdObj, , 2, 3
       rsObj.AddNew 
       rsObj("memberID")  =  Request.Form("memberID")	   
       rsObj("name")  =  Request.Form("name")
       rsObj("sex")  =  Request.Form("sex")
       rsObj("Pwd")  =  Request.Form("pwd")
       rsObj("question")  =  Request.Form("question")
       rsObj("answer")  =  Request.Form("answer")
       rsObj("email")  =  Request.Form("email")
       rsObj("phone")  =  Request.Form("phone")
       rsObj("address")  = Request.Form("address")
       rsObj("Zipcode")  =  Request.Form("Zipcode")
       rsObj.Update
       rsObj.Close
	   
       set rsObj = nothing
       CloseConn()
%>

<TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) cellPadding = 0 
                  width = "100%" borderColorLight = #aaaaaa border = 1 ID = "Table2">
  <tbody> 
  <tr bgcolor = "#FFDBBD"> 
    <td height = "26" colspan = 2 bgcolor="#FFCCCC"  > 
      <p align = "center"><b>enroll info </b></p>    </td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">ID：</div>    </td>
    <td width = "361">jw</td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">name：</div>    </td>
    <td width = "361">Zhang</td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right"><font color = "#000000">sex：</font></div>    </td>
    <td width = "361">female　 </td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">PSW：</div>    </td>
    <td width = "361">（invisible）</td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">email：</div>    </td>
    <td width = "361">111@qq.com</td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">phone：</div>    </td>
    <td width = "361">111</td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">address：</div>    </td>
    <td width = "361">qqqq</td>
  </tr>
  <tr> 
    <td width = 84 align = "right" nowrap  > 
      <div align = "right">zipcode：</div>    </td>
    <td width = "361">111</td>
  </tr>
 
  <tr bgcolor = "#FFDBBD"> 
    <td height = "26" colspan = "2" align = "right" nowrap bgcolor="#FFCCCC"  > 
      <div align = "center"> 
        <input type = "button" value = " shoppping now" name = "B4" onClick = "window.location.href = './';" style = "border: 1px solid #7D85A2;  font-size:9pt" ID = "Button1">
      </div>    </td>
  </tr>
  </tbody> 
</table></TD></TR>
</TBODY></TABLE>