<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->  
<%
if session("codusuario")<>"" then
	telefono = request("numtelefono")
	session("telefono") = telefono
end if