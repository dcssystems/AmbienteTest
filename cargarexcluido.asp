<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admexcluido.asp") then
	buscador=obtener("buscador")	
	
		if obtener("agregardato")<>"" then
		archivotxt=obtener("archivotxt")	
			%>
				<script language=javascript>
					alert("<%if archivotxt="OK" then%>Se cargó el archivo correctamente.\nSe procederá con la actualización y al finalizar se enviará un email de confirmación.<%else%>No se pudo cargar el archivo, vuelva a intentarlo.<%end if%>");
					window.close();
				</script>			
			<%				
		else
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Nueva Carga de Excluidos</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				function agregar()
				{			
					if(checkFile()==true)
					{
					document.formula.agregardato.value=1;
					document.formula.submit();
					}
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
			<script language="javascript">
			function checkFile()
			{
			    var fileElement = document.getElementById("archivotxt");
			    var fileExtension = "";
			    if (fileElement.value.lastIndexOf(".") > 0) {
			        fileExtension = fileElement.value.substring(fileElement.value.lastIndexOf(".") + 1, fileElement.value.length);
			    }
			    if (fileExtension == "txt" || fileExtension == "TXT") {
			        return true;
			    }
			    else {
			        alert("Debe seleccionar un archivo con extensión *.txt");
			        return false;
			    }
			}
			</script>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>
					<form name=formula method=post enctype="multipart/form-data" action="uploadfileexcluido.asp">
					<tr height=20>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b>Nueva Carga de Excluidos</b></b></font>
						</td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Seleccione archivo:</font></td>
						<td><input type="file" name="archivotxt" id="archivotxt" accept=".txt"><font face=Arial size=2 color=#483d8b>&nbsp;(*.txt)</font></td>
					</tr>
					<tr>
						<td width=30% ><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;</font></td>
						<td><input type=checkbox name="activo" style="font-size: xx-small;" checked>&nbsp;&nbsp;<font face=Arial size=2 color=#483d8b>Inactivar Excluidos Actuales</font></td>
					</tr>					
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codexcluido="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
						<input type=hidden name="agregardato" value="">
						<input type=hidden name="codexcluido" value="<%=codexcluido%>">
						<input type=hidden name="vistapadre" value="<%=obtener("vistapadre")%>">
						<input type=hidden name="paginapadre" value="<%=obtener("paginapadre")%>">
					</form>						
					</table>
			</body>
		</html>	
		<%		
		end if
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

