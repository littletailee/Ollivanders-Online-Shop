
 <TABLE cellSpacing = 0 cellPadding = 0 width = "100%" border = 0 ID = "Table1">
   <TBODY>
   <TR>
     <TD><!-- #include file  =  "ProductTypeList.asp"--></TD>
   </TR>
 </TBODY>
</TABLE>

<!-- #include file = "ShowSpecialProduct.asp" -->
<TABLE cellSpacing = 0  cellPadding = 0 width = "100%" border = 0 ID = "Table2">
   <TBODY>
   <TR>
     <TD>
			<%	
				call ShowSpecialProduct("Hot", 6) 
			%>
  </TD>
  </TR>
 </TBODY>
</TABLE>

<!-- #include file = "ShowSpecialProduct.asp" -->
<TABLE cellSpacing = 0  cellPadding = 0 width = "100%" border = 0 ID = "Table2">
   <TBODY>
   <TR>
     <TD>
			<%	
				call ShowSpecialProduct("Cheap", 6) 
			%>
  </TD>
  </TR>
 </TBODY>
</TABLE>

      
