<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capax.asp-->
<!--#include file=capa2.asp-->  
<% 
''if session("codusuario")<>"" then
	conectar
	
	    sql="select top 1 territorio from Reporte_ida"
	    consultar sql,RS
	    response.Write sql	
        territorio=rs.fields(0)
	    RS.Close
	    response.Write territorio	
	     

	desconectar
''end if
%>



