<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") then
		codigocentral=obtener("codigocentral")
		contrato=obtener("contrato")
		fechadatos=obtener("fechadatos")
		fechagestion=obtener("fechagestion")
	
		if obtener("agregardato")<>"" then
			emailnuevo=trim(obtener("emailnuevo"))
			
			existeemail=0
			
			if emailnuevo<>"" then
			sql="select count(*) from emailnuevo where codigocentral = '" & codigocentral & "' and  email='" & emailnuevo & "'"
			end if
			consultar sql,RS
			existeemail=RS.Fields(0)
			RS.Close		
			
				
			sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
			consultar sql,RS	
			maxfechagestion=rs.fields(0)
			RS.Close	
			
			if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
				vistabusqueda="VERULTIMPAGADO"
			else
				vistabusqueda="VERIMPAGADO"
			end if		
			
			if emailnuevo<>"" then
			sql="select count(*) from " & vistabusqueda & " where codigocentral = '" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' and email='" & emailnuevo & "'"
			end if
			consultar sql,RS
			existeemail=existeemail + RS.Fields(0)
			RS.Close		
			
			existeinactivo=0
			if existeemail > 0 then
				sql="select top 1 activo from emailnuevo where codigocentral = '" & codigocentral & "' and email='" & emailnuevo & "' order by activo"
				consultar sql,RS
				if not RS.EOF then
				    if RS.Fields(0)=0 then
					    existeinactivo=1
				    end if
				end if
				RS.Close			
			end if			
							
			if existeemail=0 or existeinactivo=1 then			
				if existeinactivo=0 then		
				sql="insert into emailnuevo(codigocentral,email,activo,usuarioregistra,fecharegistra) values ('" & codigocentral & "','" & emailnuevo & "',1," & session("codusuario") & ",getdate())"
				else
				sql="update emailnuevo set activo=1,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codigocentral = '" & codigocentral & "' and  email='" & emailnuevo & "'"
				end if
				''Response.Write sql
				conn.execute sql									
				%>
				<script language=javascript>
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre1")= "verimpagado.asp" then%>window.open("<%=obtener("paginapadre1")%>?vistapadre=<%=obtener("vistapadre")%>&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>&agregueemail=1" ,"<%=obtener("vistapadre1")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("El E-mail ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codoficina="" then%>Nuevo <%end if%>E-mail</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				function agregar()
				{

				    if (!isEmailAddress(formula.emailnuevo)) { alert("Debe ingresar un e-Mail válido."); return; }
				
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				function actualizaubigeo()
				{
				document.formula.ubigeoact.value = 1;
				document.formula.agregardato.value="";
				document.formula.submit();
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
				<form name=formula method=post action="adicionaremail.asp">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>		
					<tr bgcolor="#007DC5">
					<td colspan="2" align="left" height="18"><font size=2 face=Arial color="#FFFFFF"><b>Agregar E-mail</b></font>
					</td>
					</tr>										  
					
					
					
					<tr>
						<td width=60  bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;E-mail:</font></td>
						<td bgcolor ="#f5f5f5"><input name="emailnuevo" type=text maxlength=200 value="" style="font-size: x-small; width: 220px;"></td>
					</tr>
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
					</table>
					<input type=hidden name="agregardato" value="">
					<input type=hidden name="ubigeoact" value="">
					<input type=hidden name="codoficina" value="<%=codcliente%>">
					<input type=hidden name="vistapadre1" value="<%=obtener("vistapadre1")%>">
					<input type=hidden name="paginapadre1" value="<%=obtener("paginapadre1")%>">
					<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
					<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
					<input type="hidden" name="codigocentral" value="<%=codigocentral%>">
					<input type="hidden" name="contrato" value="<%=contrato%>">
					<input type="hidden" name="fechadatos" value="<%=fechadatos%>">
					<input type="hidden" name="fechagestion" value="<%=fechagestion%>">					
				</form>	
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

