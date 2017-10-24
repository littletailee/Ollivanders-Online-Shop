
<%
dim strSQL, rsObj

dim memberID
memberID = Request.Form("memberID")
if memberID = "" then
%>
	<script language = Javascript>
	<!--
	alert("ID can't be empty");
	window.history.go(-1);
	//-->
	                            </script>
<%
	Response.End
end if
strSQL = "SELECT memberID FROM member  WHERE memberID = '"&memberID&"'"
set rsObj = conn.execute (strSQL)
if not rsObj.eof then
%>
	<script language = Javascript>
	<!--
	alert("the user has already existed,please choose another ID");
	window.history.go(-1);
	//-->
	                            </script>
<%
	Response.End
else %><style type="text/css">
<!--
.STYLE1 {color: #333333}
-->
</style>	
<table>
<tbody>
  <TR>
    <TD height = 62 align = "center" width = "96%">
         
          <FORM language = javascript name = form1 onsubmit = "return form1_onsubmit()" method = post action = "regStep3.asp" ID = "Form1">
            <div align = "center">
              <center> 
            <TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) cellPadding = 0 
                  width = "730" borderColorLight = #AAAAAA border = 1 ID = "Table10">
              <TBODY> 
              <TR bgcolor = "#3979C6"> 
                <TD colSpan = 2 bgcolor = "#FFCCCC" width = "726"><B>Enroll info£º</B></TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">ID£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626">
                  <input type = text value = "<%=memberID%>" name = "memberID" size = "12" ReadOnly style = "font-size: 9pt; border: 1px solid #7D85A2;  " ID = "Text1"></TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">name£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"> 
                  <INPUT maxLength = 20 size = 10 name = name style = "font-size: 9pt; border: 1px solid #7D85A2;  " ID = "Text2"></TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right"><FONT color = #000000>sex£º</FONT></div>
                </TD>
                <TD bgColor = #ffffff width = "626"><FONT color = #000000> 
                  <INPUT id = sex  
            type = radio CHECKED value = male name = sex>
                  male 
                  <INPUT id = "Radio1" type = radio  
            value = female name = sex>
                  female</FONT> </TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">PSW£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"> 
                  <P><FONT color = #000000> 
                    <INPUT type = password maxLength = 18 name = pwd size = "20" style = "font-size: 9pt; border: 1px solid #7D85A2;  " ID = "Password1">
                    </FONT>4~16 characters </P>
                </TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">confirm psw£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"><FONT color = #000000> 
                  <INPUT  
            type = password maxLength = 18 name = PasswordConfirm size = "20" style = "font-size: 9pt; border: 1px solid #7D85A2;  " ID = "Password2">
                </FONT></TD>
              </TR>
              
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">Email£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"> 
                  <INPUT maxLength = 40 name = email size = "20" style = "font-size: 9pt; border: 1px solid #7D85A2;  " ID = "Text5"></TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">Phone£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"> 
                  <INPUT id = "phone" maxLength = 30   
            name = phone size = "20" style = "font-size: 9pt; border: 1px solid #7D85A2;  "></TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">Address£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"> 
                  <INPUT maxLength = 200 size = 60 name = "address" style = "font-size: 9pt; border: 1px solid #7D85A2;  " ></TD>
              </TR>
              <TR> 
                <TD width = 98 bgColor = #ffffff nowrap> 
                  <div align = "right">Zip chode£º</div>
                </TD>
                <TD bgColor = #ffffff width = "626"> 
                  <INPUT maxLength = 15 name = "zipcode" size = "20" style = "font-size: 9pt; border: 1px solid #7D85A2;  " >
                </TD>
              </TR>
              <TR> 
                <TD bgColor = #ffffff colSpan = 2 width = "726"></TD>
              </TR>
              <TR id = tr_51go style = "DISPLAY: none" bgColor = #ffffff> 
                <TD colSpan = 2 width = "726">¡¡</TD>
              </TR>
              </TBODY>
            </TABLE>
            
              </center>
            </div>
            <INPUT  type = submit value = "submit" name = button1 style = "font-size: 9pt; border: 1px solid #7D85A2; " ID = "Submit1">
            <INPUT  id = button2 type = reset value = "reset" name = reset style = "font-size: 9pt; border: 1px solid #7D85A2; ">   
      <BR></FORM></DIV></TD></TR></TBODY></TABLE>
<%end if %>