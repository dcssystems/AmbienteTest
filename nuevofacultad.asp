<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admfacultad.asp") then
	buscador=obtener("buscador")	
	codfacultad=obtener("codfacultad")
		if obtener("agregardato")<>"" then
		codgrupofacultad=obtener("codgrupofacultad")
		descripcion=obtener("descripcion")
		pagina=obtener("pagina")
		orden=obtener("orden")
		if not isNumeric(orden) then
			orden="0"
		end if						
										
									
			existefacultad=0
			
			if codfacultad<>"" then
			sql="select count(*) from facultad where descripcion='" & descripcion & "' and codfacultad<>" & codfacultad & " and codgrupofacultad=" & codgrupofacultad 
			else
			sql="select count(*) from facultad where descripcion='" & descripcion & "' and codgrupofacultad=" & codgrupofacultad 
			end if
			consultar sql,RS
			existefacultad=RS.Fields(0)
			RS.Close			
			if existefacultad=0 then			
				if obtener("agregardato")="1" then		
				sql="insert into facultad (codgrupofacultad,descripcion,pagina,orden,usuarioregistra,fecharegistra) values (" & codgrupofacultad & ",'" & descripcion & "','" & pagina & "','" & orden & "'," & session("codusuario") & ",getdate())"
				else
					sql="update facultad set codgrupofacultad=" & codgrupofacultad & ",descripcion='" & descripcion & "',pagina='" & pagina & "',orden=" & orden & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codfacultad=" & codfacultad
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
					<%if obtener("paginapadre")="admfacultad.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("El usuario ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codfacultad<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from facultad A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where a.codfacultad = " & codfacultad
					consultar sql,RS
					descripcion=rs.Fields("descripcion")
					codgrupofacultad=rs.Fields("codgrupofacultad")		
					pagina=rs.Fields("pagina")		
					orden=rs.Fields("orden")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codfacultad="" then%>Nueva <%end if%>Facultad</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				<%if codfacultad="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(trim(formula.pagina.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(trim(formula.pagina.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
					
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
						<form name=formula method=post action="nuevofacultad.asp">					
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codfacultad="" then%>Nueva <%end if%>Facultad</b></b></font>
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
						<td width=20%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descripción:</font></td>
						<td><input name="descripcion" type=text maxlength=200 value="<%=Descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
					<td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Grupo:</font></td>
					<td bgcolor="#f5f5f5">
						<select name="codgrupofacultad" style="font-size: xx-small; width: 200px;">
						<%
						sql = "select codgrupofacultad, descripcion from grupofacultad order by orden"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codgrupofacultad")%>" <% if codgrupofacultad<>"" then%><% if RS.fields("codgrupofacultad")=int(codgrupofacultad) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>
						</select>
						</td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Link:</font></td>
						<td><input name="pagina" type=text maxlength=200 value="<%=pagina%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Orden:</font></td>
						<td bgcolor="#f5f5f5"><input name="orden" type=text maxlength=50 value="<%=orden%>" style="font-size: xx-small; width: 60px; text-align: right"></td>
					</tr>			
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codfacultad="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>

							<input type=hidden name="agregardato" value="">
							<input type=hidden name="codfacultad" value="<%=codfacultad%>">
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

