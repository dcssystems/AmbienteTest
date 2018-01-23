<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("archivosdistribucion.asp") then
	buscador=obtener("buscador")	
	
	sql="select a.usuario, a.administrador, b.codagencia,b.razonsocial from usuario a left join agencia b on a.codagencia = b.codagencia where codusuario = " & session("codusuario")
	consultar sql,RS	
	usuario=rs.fields("usuario")
	fadmin=rs.fields("administrador")
	agencia=iif(isNull(rs.Fields("razonsocial")),"Sin asignar",rs.Fields("razonsocial"))
	codagencia=rs.fields("codagencia")
	rs.close
	
	tiempoexport=Now()
	
	if codagencia<>"" then
		''sql="select convert(varchar,max(fecharegistra),103), convert(varchar,max(fechagestion),103), convert(varchar,max(fecharegistra),112) from DistribucionDiaria where codagencia=" & codagencia
		sql="select convert(varchar,(select top 1 fecharegistra from DistribucionDiaria where codagencia=" & codagencia & " and FechaGestion=max(A.fechagestion)),103),convert(varchar,max(fechagestion),103),convert(varchar,(select top 1 fecharegistra from DistribucionDiaria where codagencia=" & codagencia & " and FechaGestion=max(A.fechagestion)),112) from BusquedaDistribucion A where codagencia=" & codagencia
		consultar sql,RS
		ultfecha=RS.Fields(0)
		fechagestion=RS.Fields(1)
		archfecha=RS.Fields(2)
		rs.close
	end if
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Archivos Distribución</title>
			<script language=javascript>
				function descargar()
				{		
					<%
					sql="select descripcion,valortexto1 from parametro where descripcion='RutaWebExportar'"
					consultar sql,RS
					RutaWebExportar=RS.Fields(1)
					RS.Close					
					%>	
					window.open("<%=RutaWebExportar%>/<%=archfecha%>_<%=mid(agencia,1,2)%>_<%=codagencia%>.txt?time=<%=tiempoexport%>","_self");
					//window.open("descargararchivo.asp?exportados=1&nomarch=<%=archfecha%>_<%=UCase(mid(agencia,1,2))%>_<%=codagencia%>.txt","_self");
				}
				function trim(string)
				{
					while(string.substr(0,1)==" ")
					string = string.substring(1,string.length) ;
					while(string.substr(string.length-1,1)==" ")
					string = string.substring(0,string.length-2) ;
					return string;
				}			
				function isEmailAddress(theElement)
				{
				var s = theElement.value;
				var filter=/^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ig;
				if (s.length == 0 ) return true;
				   if (filter.test(s))
				      return true;
				   else
					theElement.focus();
					return false;
				}					
			</script>
			<style>
			A {
				FONT-SIZE: 12px; COLOR: #483d8b; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
			}
			A:visited {
				TEXT-DECORATION: none; COLOR: #483d8b;
			}
			A:hover {
				COLOR: #483d8b; FONT-FACE:"Arial"; TEXT-DECORATION: none
			}			
			</style>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=30%>
					<form name=formula method=post enctype="multipart/form-data" action="uploadfilerespuesta.asp">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;Descarga de Archivos de Distribución</b></font>
						</td>
					</tr>
					<!--<tr>
						<td width=15%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Usuario:</font></td>
						<td width=25%><font face=Arial size=2 color=#483d8b>&nbsp;<%=usuario%></font></td>
						<td rowspan=4 valign=bottom><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a></td>
					</tr>-->
					<tr>
						<td width=15%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Última&nbsp;Distribución&nbsp;Disponible:</font></td>
						<td><font face=Arial size=2 color=#483d8b><b>&nbsp;<%if ultfecha<>"" then%><%=ultfecha%><%else%>No&nbsp;disponible<%end if%></b></font></td>
					</tr>
					<tr>
						<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Fecha de Gestión:</font></td>
						<td><font face=Arial size=2 color=#483d8b><b>&nbsp;<%if fechagestion<>"" then%><%=fechagestion%><%else%>No&nbsp;disponible<%end if%></b></font></td>
					</tr>
					<tr>
						<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Agencia:</font></td>
						<td><font face=Arial size=2 color=#483d8b><b>&nbsp;<%=agencia%></b></font></td>
					</tr>
					<%if ultfecha<>"" then%>
					<tr>
						<td valign=top><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descargar Archivo:</font></td>
						<td valign=top><a href="<%=RutaWebExportar%>/<%=archfecha%>_<%=mid(agencia,1,2)%>_<%=codagencia%>.rar?time=<%=tiempoexport%>" target="Arch"><!--"javascript:descargar();">--><img src="imagenes/descargar.gif" border=0 alt="Descargar Archivo" title="Descargar Archivo"></a></td>
					</tr>					
					<%end if%>					
					<tr>					
						<td colspan=2 bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>								
					</tr>
					</form>						
					</table>
			</body>
		</html>	
		<%		
		''end if
	else
	%>
	<script language="javascript">
		alert("Ud. No tiene autorización para este proceso.");
		window.open("userexpira.asp","_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	//window.open("index.html","_top");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>

