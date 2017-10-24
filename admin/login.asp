<!-- #include file = "include/conndb.asp" -->

<!-- #include file = "include/adminbase.asp" -->
<%
	dim strUserId, strPwd
	strUserId  =  RealString(Request.Form("userID"))
	strPwd  =  RealString(Request.Form("password"))
	if strUserId <> "" And strPwd <> "" then
		call CheckAdminLogin(strUserId, strPwd)
	end if	
%>
  
<HTML><HEAD>
<script language = "javascript">
	<!-- 
		if (self !=  top)
		{
			top.location = self.location;
			alert("you're offline, please login again");
		} 
	-->
	</script> 

<!-- #include file = "include/head.asp"-->

</HEAD>
<BODY>

	<TABLE align = center  cellSpacing = 0 cellPadding = 0 width = "750" border = 0 ID = "tabMain" class = tabframe>
	<TBODY><TR>
		<TD >
			<% ShowHead() %>
		</TD>
	</TR></TBODY></TABLE>
	
	<TABLE align = center cellSpacing = 0 cellPadding = 0 width = "750" border = 0 height = "211" ID = "Table3"  class = tabframe>
	<TBODY><TR>
		<TD vAlign = middle align = center >
			<% call ShowAdminLogin() %>
		</TD>
	</TR></TBODY></TABLE>
	<hr size="1" width="50%" align="center">
</BODY>
</HTML>

