<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admagencia.asp") then
	buscador=obtener("buscador")	
	codagencia=obtener("codagencia")
		if obtener("agregardato")<>"" then
		codgrupoagencia=obtener("codgrupoagencia")
		telefono=obtener("telefono")
		direccion=obtener("direccion")
		razonsocial=obtener("razonsocial")
		plaza=obtener("plaza")
		grupoefectividad=obtener("grupoefectividad")
		subgrupoefectividad=obtener("subgrupoefectividad")
		nombresunat=obtener("nombresunat")
		rucsunat=obtener("rucsunat")
		if obtener("activo")<>"" then activo="1" else activo="0" end if	
						
			existeagencia=0
			
			if codagencia="" then
				sql="select count(*) from agencia where razonsocial='" & razonsocial & "'"
			else
				sql="select count(*) from agencia where razonsocial='" & razonsocial & "' and codagencia<>" & codagencia
			end if			
			consultar sql,RS
			existeagencia=RS.Fields(0)
			RS.Close			
			if existeagencia=0 then			
				if obtener("agregardato")="1" then	
					if codgrupoagencia="" then
						sql="insert into agencia (codgrupoagencia,razonsocial,plaza,grupoefectividad,subgrupoefectividad,nombresunat,rucsunat,telefono,direccion,usuarioregistra,fecharegistra,activo) values (null,'" & razonsocial & "','" & plaza & "','" & grupoefectividad & "','" & subgrupoefectividad & "','" & nombresunat & "','" & rucsunat & "','" & telefono & "','" & direccion & "'," & session("codusuario") & ",getdate()," & activo & " )"
					else
						sql="insert into agencia (codgrupoagencia,razonsocial,plaza,grupoefectividad,subgrupoefectividad,nombresunat,rucsunat,telefono,direccion,usuarioregistra,fecharegistra,activo) values (" & codgrupoagencia & ",'" & razonsocial & "','" & plaza & "','" & grupoefectividad & "','" & subgrupoefectividad & "','" & nombresunat & "','" & rucsunat & "','" & telefono & "','" & direccion & "'," & session("codusuario") & ",getdate()," & activo & " )"
					end if
				else
					if codgrupoagencia="" then
						sql="update agencia set codgrupoagencia=null,razonsocial='" & razonsocial & "',plaza='" & plaza & "',grupoefectividad='" & grupoefectividad & "',subgrupoefectividad='" & subgrupoefectividad & "',nombreSunat='" & nombresunat & "',rucsunat='" & rucsunat & "',telefono='" & telefono & "',direccion='" & direccion & "',usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate(), activo=" & activo & " where codagencia=" & codagencia
					else
						sql="update agencia set codgrupoagencia=" & codgrupoagencia & ",razonsocial='" & razonsocial & "',plaza='" & plaza & "',grupoefectividad='" & grupoefectividad & "',subgrupoefectividad='" & subgrupoefectividad & "',nombreSunat='" & nombresunat & "',rucsunat='" & rucsunat & "',telefono='" & telefono & "',direccion='" & direccion & "',usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate(), activo=" & activo & " where codagencia=" & codagencia
					end if
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
					<%if obtener("paginapadre")="admagencia.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("La agencia ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codagencia<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from agencia A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where A.codagencia = " & codagencia
					consultar sql,RS
					codgrupoagencia=rs.Fields("codgrupoagencia")		
					razonsocial=rs.Fields("razonsocial")
					telefono=rs.Fields("telefono")
					nombresunat=rs.Fields("nombreSunat")
					rucsunat=rs.Fields("rucSunat")
					direccion=rs.Fields("direccion")
					plaza=rs.Fields("plaza")
					grupoefectividad=rs.Fields("grupoefectividad")
					subgrupoefectividad=rs.Fields("subgrupoefectividad")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					activo=rs.Fields("activo")
					RS.Close
			else
				activo="1"
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codagencia="" then%>Nueva <%end if%>Agencia</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				<%if codagencia="" then%>
				function agregar()
				{
					if(trim(formula.razonsocial.value)==""){alert("Debe ingresar una Razon Social para ingreso al Sistema.");return;}
					if(trim(formula.direccion.value)==""){alert("Debe ingresar una Dirección para ingreso al Sistema.");return;}													
					if(trim(formula.telefono.value)==""){alert("Debe ingresar un Teléfono para ingreso al Sistema.");return;}	
					
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.razonsocial.value)==""){alert("Debe ingresar una Razon Social para ingreso al Sistema.");return;}
					if(trim(formula.direccion.value)==""){alert("Debe ingresar una Dirección para ingreso al Sistema.");return;}
					if(trim(formula.telefono.value)==""){alert("Debe ingresar un Teléfono para ingreso al Sistema.");return;}		
					
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
					<form name=formula method=post action="nuevoagencia.asp">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codagencia="" then%>Nueva <%end if%>Agencia</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr height=20>
						<td colspan=2 align=right><font face=Arial size=1 color=#483d8b>Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<%if codagencia = "" then%>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código:</font></td>
						<td>
							<font face=Arial size=2 color=#483d8b>Nuevo</font></td>
					</tr>
					
					<%else%>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código:</font></td>
						<td>
							  
							  <input type=hidden name="codagencia" value="<%=codagencia%>">
							  <font face=Arial size=2 color=#483d8b><%=codagencia%></font></td>
						
					</tr>
					<%end if%>
					<tr>
					<td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Grupo Agencia:</font></td>
					<td bgcolor="#f5f5f5">
						<select name = "codgrupoagencia" style="font-size: xx-small; width: 200px;">
						<option value="">Seleccionar Grupo</option>
						<%
						sql = "select codgrupoagencia, descripcion from grupoagencia order by codgrupoagencia"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codgrupoagencia")%>" <% if codgrupoagencia<>"" then%><% if RS.fields("codgrupoagencia")=int(codgrupoagencia) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>
						</select>
						</td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Nombre:</font></td>
						<td><input name="razonsocial" type=text maxlength=200 value="<%=razonsocial%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Plaza:</font></td>
						<td bgcolor="#f5f5f5">
						    <select name = "plaza" style="font-size: xx-small; width: 200px;">
						        <%
						        sql = "select codplaza, descripcion from Plaza order by codplaza"
						        consultar sql,RS
						        Do While Not  RS.EOF
						        %>
						        <option value="<%=RS.Fields("descripcion")%>" <% if plaza<>"" then%><% if RS.fields("descripcion")=plaza then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						        <%
						        RS.MoveNext
						        loop
						        RS.Close
						        %>
						</td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Grupo efectividad:</font></td>
						<td><input name="grupoefectividad" type=text maxlength=200 value="<%=grupoefectividad%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Sub-Grupo efectividad:</font></td>
						<td bgcolor="#f5f5f5"><input name="subgrupoefectividad" type=text maxlength=200 value="<%=subgrupoefectividad%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>					
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Razon Social:</font></td>
						<td bgcolor ="#f5f5f5"><input name="nombresunat" type=text maxlength=200 value="<%=nombresunat%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;RUC:</font></td>
						<td bgcolor="#f5f5f5"><input name="rucsunat" type=text maxlength=200 value="<%=rucsunat%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Teléfono:</font></td>
						<td><input name="telefono" type=text maxlength=50 value="<%=telefono%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Dirección:</font></td>
						<td bgcolor="#f5f5f5"><input name="direccion" type=text maxlength=200 value="<%=direccion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Estado:</font></td>
						<td><input type=checkbox name="activo" style="font-size: xx-small;" <%if activo=1 then%> checked<%end if%>>&nbsp;&nbsp;<font face=Arial size=2 color=#483d8b>Activo</font></td>
					</tr>			
					<tr>					
						<td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#f5f5f5" align=right height=40><%if codagencia="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
						<input type=hidden name="agregardato" value="">
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

