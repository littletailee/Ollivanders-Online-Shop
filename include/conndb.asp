
<%
  Option Explicit	

  dim conn		 
  dim connstr		
  dim db			  

  db = "Database/shop.mdb"	
  connstr = "Provider = Microsoft.Jet.OLEDB.4.0;Data Source = " & Server.MapPath(db)
  
  Set conn  =  Server.CreateObject("ADODB.Connection")
  if err Then
	  err.clear
  end if

  conn.Open connstr

  sub CloseConn()
	  conn.Close()
	  Set conn  =  Nothing
  end sub

  function RealString(strSrc)
	  RealString  =  Replace(Trim(strSrc), "'", "¡¯")
  end function

  function Convert(strSrc)
	  Convert  =  Server.HTMLEncode(Replace(Trim(strSrc), "'", "¡¯"))
	  Convert  =  Replace(Convert, chr(13), "<br>")
  end function

%>

