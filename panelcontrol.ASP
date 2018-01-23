<%@ LANGUAGE = VBScript.Encode %>
<% 
if session("codusuario")<>"" then
%>
	<title>Sistema Web de Gestión de Cobranzas BBVA</title>
	<frameset rows="*,0" frameborder=0>
		<!--<frame src="herramientas.asp" name="herramientas" scrolling=no noresize>-->
		<!--<frame src="bienvenida.html" name="sistema">-->	
		<frame src="herramientas.asp" name="sistema" scrolling=auto noresize>
	</frameset>
<%
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	window.open("index.html","_top");
</script>
<%
end if
%>



