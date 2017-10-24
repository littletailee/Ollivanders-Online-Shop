<%

	function TypeToSelect(nTypeId, szCtlName)
		dim strSQL
		dim rsType
		dim tempi
		dim nID, szName
		Response.Write	"<SELECT name =""" & szCtlName & """ ID=""Select1"">"
		Response.Write	"<option value = ""0"">all categories</option>"
		
		strSQL = "SELECT * FROM ProductType "
		set rsType = Server.CreateObject("ADODB.RecordSet")
		rsType.Open strSQL, conn, 2,1
		if rsType.EOF And rsType.BOF then
			Response.Write	"</SELECT>"
			rsType.Close()
			Set rsType = Nothing
			exit function
		else
			do while not rsType.eof 
				nID = rsType("id")
				szName = rsType("name")
				Response.Write "<option value = "& nID
				
				if CInt(nID) = CInt(nTypeId) then Response.Write " selected"
				Response.Write ">"
				Response.Write szName &"</option>"
				rsType.MoveNext
			loop
		end if
		Response.Write	"</SELECT>"
		rsType.Close()
		Set rsType = Nothing
		
	end function
%>	