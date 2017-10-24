<!-- #include file = "include/conndb.asp"-->
<!-- #include file = "include/config.asp"-->
<!-- #include file = "include/memberbase.asp"-->
<!-- #include file = "include/productbase.asp"-->
<!-- #include file = "include/cartbase.asp"-->


<HTML><HEAD>
<!-- #include file = "include/head.asp"-->
</HEAD>
<BODY>


<TABLE align = center  cellSpacing = 0 cellPadding = 0 width = "1144" border = 0 ID = "tabMain" class = tabframe>
	<TBODY>
	<TR>
		<TD >
			<% ShowHeadAndMenu() %>
		</TD>
	</TR>
	</TBODY>
</TABLE>


<TABLE align = center cellSpacing = 0 cellPadding = 0 width = "1144" border = 0 height = "411" ID = "Table3"  class = tabframe>
<TBODY><TR>

	<TD vAlign = top align = middle width = 300 height = "411">
		<table width = 95% ID = "Table1"><tr><td>
		<!-- #include file = "include/nav.asp"-->
		</td></tr></table>
	</TD>
	
	<TD vAlign = top align = right height = "411">
		<TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) 
		cellPadding = 0 width = "100%" bgColor = #666666 
		borderColorLight = #aaaaaa border = 1 ID = "Table2">
		<TBODY><TR>
			<TD>
			<TABLE cellSpacing = 0 cellPadding = 0 width = "100%" bgColor = #ffffff border = "0" style = "border-collapse: collapse" bordercolor = "#111111" ID = "Table5">
				<TBODY>
				<TR>
				<TD >
					<!-- #include file = "orderInfo.asp"-->
				</TD>
				</TR></TBODY></table>                                
			</TD>
		</TR></TBODY></TABLE>
		
	</TD>
</TR></TBODY></TABLE>
	
	
    <!-- #include file = "include/foot.asp"-->
	</BODY>
</HTML>
