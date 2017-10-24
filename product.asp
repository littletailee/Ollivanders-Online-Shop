<!-- #include file = "include/sysbase.asp"-->
<!-- #include file = "include/ProductSearchForm.asp"-->

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
	<%
			dim Keyword
			dim ProductType
			dim page
			Keyword  =  RealString(Request.Form("Keyword"))
			ProductType  =  RealString(Request("ProductType"))
			page  =  RealString(Request.QueryString("page"))
			if ProductType  =  "" then ProductType  =  0
	
	%>
	
		<TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) 
			cellPadding = 0 width = "100%" bgColor = #666666 
			borderColorLight = #aaaaaa border = 1 ID = "Table18">
		<TBODY><TR>
			<TD>
				<TABLE cellSpacing = 0 cellPadding = 0 width = "100%" bgColor = #ffffff border = "0" 
						style = "border-collapse: collapse" bordercolor = "#111111" ID = "Table19">
				<TBODY>
				<TR>
					<TD bgColor = #f0f8ff><% call ShowSearchForm(Keyword, ProductType) %></TD>
				</TR></TBODY></TABLE>
		</TR></TBODY></TABLE>
		<%
			if Keyword <> "" then
				Response.Write "<div align = left>search for <font color = red><b>" & Keyword & "</b></font> </div>"
			end if
		%>
		
		<TABLE cellSpacing = 0 borderColorDark = rgb(210,232,255) 
		cellPadding = 0 width = "100%" bgColor = #666666 
		borderColorLight = #aaaaaa border = 1 ID = "Table2">
		<TBODY><TR>
			<TD>
			<TABLE cellSpacing = 0 cellPadding = 0 width = "100%" bgColor = #ffffff border = "0" style = "border-collapse: collapse" bordercolor = "#111111" ID = "Table5">
				<TBODY>
				<TR>
				<TD ><% 
					call ShowNewProductPreview("All", ProductType, Keyword, page) %></TD>
				</TR></TBODY></table>                                
			</TD>
		</TR></TBODY></TABLE>
		
	</TD>
	
</TR></TBODY></TABLE>

	
	
    <!-- #include file = "include/foot.asp"-->
	</BODY>
</HTML>
