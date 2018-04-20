<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
usr=obtener("usr")
pwd=obtener("pwd")

conectar
if validarusuario(usr,pwd,mensajevalida) then
	
	Response.Redirect "dcs_herramientas.asp"
else
%>
<script language="javascript">
	alert("<%=mensajevalida%>");
	history.back();
</script>
<%
end if
desconectar
%>

