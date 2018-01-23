<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admmarcacliente.asp") then
	buscador=obtener("buscador")	
	codmarcacliente=obtener("codmarcacliente")
		if obtener("agregardato")<>"" then
		codmarca=obtener("codmarca")
		codigocentral=obtener("codigocentral")
		if obtener("activo")<>"" then activo="1" else activo="0" end if	
										
			codmarcaclientenuevo=obtener("codmarcaclientenuevo")
						
			existemarcacliente=0
			
			if codmarcaclientenuevo<>"" then
				sql="select count(*) from marcacliente where codmarcacliente = '" & codmarcacliente & "' or codigocentral='" & codigocentral& "'"
			else			
				sql="select count(*) from marcacliente where codigocentral='" & codigocentral & "' and codmarcacliente<>'" & codmarcacliente & "'"
			end if
			consultar sql,RS
			existemarcacliente=RS.Fields(0)
			RS.Close			
			if existemarcacliente=0 then			
				if obtener("agregardato")="1" then		
					sql="insert into marcacliente (codmarca,codigocentral,activo,usuarioregistra,fecharegistra) values (" & codmarca & ",'" & codigocentral & "'," & activo & "," & session("codusuario") & ",getdate())"
				else
					sql="update marcacliente set codmarca=" & codmarca & ",codigocentral='" & codigocentral & "',activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codmarcacliente=" & codmarcacliente
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
					<%if obtener("paginapadre")="admmarcacliente.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("El marcacliente ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codmarcacliente<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from marcacliente A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where A.codmarcacliente = " & codmarcacliente
					consultar sql,RS
					codigocentral=rs.Fields("codigocentral")
					codmarca=rs.Fields("codmarca")		
					activo=rs.Fields("activo")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			else
				activo="1"
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codmarcacliente="" then%>Nuevo <%end if%>marcacliente</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				<%if codmarcacliente="" then%>
				function agregar()
				{
					if(trim(formula.codigocentral.value)==""){alert("Debe ingresar un Cliente para ingreso al Sistema.");return;}
																	
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.codigocentral.value)==""){alert("Debe ingresar un Cliente para ingreso al Sistema.");return;}
					
					document.formula.agregardato.value=2;
					document.formula.submit();
				}				
				<%end if%>
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
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>
					<form name=formula method=post action="nuevomarcacliente.asp">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codmarcacliente="" then%>Nuevo <%end if%>Marca Cliente</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr height=20>
						<td colspan=2 align=right><font face=Arial size=1 color=#483d8b>Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código:</font></td>
						<td><%if codmarcacliente = "" then%>
							  <font face=Arial size=2 color=#483d8b>Nuevo</font>
							  <input type=hidden name="codmarcaclientenuevo" value="<%=codmarcacliente%>"></td>
							   <%else%>
							  <font face=Arial size=2 color=#483d8b><b><%=codmarcacliente%></b></font></td>
							  <%end if%>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código Cliente:</font></td>
						<td><input name="codigocentral" type=text maxlength=200 value="<%=codigocentral%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
					<td bgcolor ="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Marca:</font></td>
					<td bgcolor ="#f5f5f5">
						<select name = "codmarca" style="font-size: xx-small; width: 200px;">
						<%
						sql = "select codmarca, descripcion from marca order by codmarca"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codmarca")%>" <% if codmarca<>"" then%><% if RS.fields("codmarca")=int(codmarca) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>
						</select>
						</td>
					</tr>
					<tr>
						<td bgcolor ="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Estado:</font></td>
						<td bgcolor ="#f5f5f5"><input type=checkbox name="activo" style="font-size: xx-small;" <%if activo=1 then%> checked<%end if%>>&nbsp;&nbsp;<font face=Arial size=2 color=#483d8b>Activo</font></td>
					</tr>			
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codmarcacliente="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
						<input type=hidden name="agregardato" value="">
						<input type=hidden name="vistapadre" value="<%=obtener("vistapadre")%>">
						<input type=hidden name="paginapadre" value="<%=obtener("paginapadre")%>">
						<input type=hidden name="codmarcacliente" value="<%=codmarcacliente%>">
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

