
<!-- #include file = "config.asp" -->

<TITLE>Ollivanders</TITLE>

<LINK href = "include/main.css" type = text/css rel = stylesheet>
<script language = Javascript src = "include/common.js"></script>


<%
function ShowHead()
%>
	
	<style type="text/css">
<!--
.STYLE1 {font-family: "Courier New", Courier, monospace}
-->
    </style>
	<hr size="1">
<%	
end function


function ShowHeadAndMenu()
	
	call ShowHead()
%>
	<style type="text/css">
<!--
.STYLE2 {color: #7DACFC}
-->
    </style>
    <TABLE cellSpacing = "0" cellPadding = "0" width = "100%"  border = "0" ID = "Table1">
		<TBODY>
		<TR>
			<TD align = left width = 409 height = 64>&nbsp;
			  <h1><span class="STYLE2">Olliva<span class="STYLE1"></span>nders</span></h1></TD>
			<TD width="566" height = 64 align = "right"> </TD>
		</TR>
		</TBODY>
	</TABLE>
	<hr size="1">
	<table border = "0" cellspacing = "0" cellpadding = 0 width = "100%"  height = "20" ID = "Table3">
    <tr>
        <td bgcolor="#00CC33"><td width = "550" height = 30 bgcolor = "#FFFFFF" align = "left" style = "word-spacing: 5" height = "18" class = "jj1" >
        <a href = manuList.asp class = "jj1" target = "frmMain">category management
        </a>&nbsp|&nbsp<a href = productList.asp class = "jj1" target = "frmMain">product management
        </a>&nbsp|&nbsp<a href = orderList.asp class = "jj1" target = "frmMain">order management 
        </a>
        </td>
        <td width = 28></td>
        <td >
        &nbsp&nbsp<a href = "logout.asp" class = "jj1" >logout</a>
        </td>
        
     </tr>
     </table>	
	<TABLE cellSpacing = "0" cellPadding = "0"  width = 100% border = "0" height = "1" ID = "Table4">
			<TR>
				<TD height = "1" ></TD>
			</TR>
	</TABLE>
<%
end function
%>