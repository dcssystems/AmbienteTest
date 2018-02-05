<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admusuario.asp") then
		buscador=obtener("buscador")	
		codusuario=obtener("codusuario")
		if obtener("agregardato")<>"" then
			usuario=obtener("usuario")
			clave=encriptar(obtener("clave"))
			apepat=obtener("apepat")
			apemat=obtener("apemat")
			nombres=obtener("nombres")
			correo=obtener("correo")
			codtipousuario=obtener("codtipousuario")
			codagencia=obtener("codagencia")
			if codagencia="" then
				act_codagencia="null"
			else
				act_codagencia=codagencia
			end if
			codoficina=obtener("codoficina")
			if codoficina="" then
				act_codoficina="null"
			else
				act_codoficina="'" & codoficina & "'"
			end if			
			
			if obtener("fbloq")<>"" then
				fbloq="3"
			else
				fbloq="0"
			end if
			
			if obtener("administrador")<>"" then administrador="1" else administrador="0" end if
			if obtener("activo")<>"" then activo="1" else activo="0" end if
			
			existeusuario=0
			
			if codusuario<>"" then
			sql="SELECT COUNT(*) FROM Usuario WHERE usuario='" & usuario & "' AND codusuario<>" & codusuario
			else
			sql="SELECT COUNT(*) FROM Usuario WHERE usuario='" & usuario & "'"
			end if
			consultar sql,RS
			existeusuario=RS.Fields(0)
			RS.Close			
			
			if existeusuario=0 then			
				if obtener("agregardato")="1" then		
				sql="INSERT INTO usuario (usuario,clave,apepaterno,apematerno,nombres,correo,flagbloqueo,administrador,activo,usuarioregistra,fecharegistra,codtipousuario,codagencia,codoficina) VALUES ('" & usuario & "','" & clave & "','" & apepat & "','" & apemat & "','" & nombres & "','" & correo & "'," & fbloq & "," & administrador & "," & activo & "," & session("codusuario") & ",getdate()," & codtipousuario & "," & act_codagencia & "," & act_codoficina & ")"
				else
					if obtener("hclave")=obtener("clave")then
					sql="UPDATE usuario SET usuario='" & usuario & "',apepaterno='" & apepat & "',apematerno='" & apemat & "',nombres='" & nombres & "',correo='" & correo & "',flagbloqueo=" & fbloq & ",administrador=" & administrador & ",codtipousuario=" & codtipousuario & ",codagencia=" & act_codagencia & ",codoficina=" & act_codoficina & ",activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() WHERE codusuario=" & codusuario
					else
					sql="UPDATE usuario SET usuario='" & usuario & "',clave='" & clave & "',apepaterno='" & apepat & "',apematerno='" & apemat & "',nombres='" & nombres & "',correo='" & correo & "',flagbloqueo=" & fbloq & ",administrador=" & administrador & ",codtipousuario=" & codtipousuario & ",codagencia=" & act_codagencia & ",codoficina=" & act_codoficina & ",activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() WHERE codusuario=" & codusuario
					end if				
				end if
				''Response.Write sql
				conn.execute sql
				
				if codusuario="" then
					sql="SELECT codusuario FROM Usuario WHERE usuario='" & usuario & "'"
					consultar sql,RS
					codusuario=RS.Fields(0)
					RS.Close	
				end if
				
				''aquí se elimina y se agregan las facultades seleccionadas							
				sql="SELECT A.codperfil, CASE WHEN B.activo IS NULL THEN 0 ELSE B.ACTIVO END AS activo, A.descripcion, codusuario FROM perfil A LEFT OUTER JOIN usuarioperfil B ON A.codperfil = B.codperfil AND B.codusuario = " & codusuario  & " ORDER BY A.codperfil"
				consultar sql,RS
				Do While not RS.EOF
					if obtener("codperf" & RS.Fields("codperfil"))<>"" then
					enviaactivar=1					
					else
					enviaactivar=0
					end if
										
					if IsNull(RS.Fields("codusuario")) then
					'' esto es que no existe
						if enviaactivar=1 then
						''esto es que lo activo y como no existe lo inserto
						sql = "insert usuarioperfil (codusuario,codperfil,usuarioregistra, fecharegistra,activo) values (" & codusuario & "," & RS.Fields("codperfil") & "," & session("codusuario") & ",getdate()," & enviaactivar & ")"
						''Response.Write sql
						conn.execute sql
						end if
					else
					'' esto es que ya existe 
						sql = "update usuarioperfil set usuariomodifica= " & session("codusuario") & ", fechamodifica= getdate(), activo = " & enviaactivar & " where  codusuario ="& codusuario &" and codperfil ="& RS.Fields("codperfil")
						''Response.Write sql
						conn.execute sql
					end if
					
					
				RS.MoveNext
				Loop
				RS.Close
				%>
				<script language="javascript">
					<% if obtener("agregardato")="1" then %>
					//alert("Se agregó el usuario correctamente.");
					<% else %>
					//alert("Se modificó el usuario correctamente.");
					<% end if %>				
					<% if obtener("paginapadre")="dcs_admusuario.asp" then %> 
					window.open('<%=obtener("paginapadre")%>','<%=obtener("vistapadre")%>');
					<% end if %>
					window.parent.close();
				</script>			
				<%
			else
			%>
				<script language="javascript">
					alert("Error, el usuario ingresado ya existe.");
					//history.back();
					//window.close();
				</script>			
			<%				
			end if
		else
			if codusuario<>"" then
				sql="SELECT A.*,E.CodPerfil AS codtipousuario, " &_
				"B.nombres as Nombreusureg, " &_
				"B.apepaterno as Apepatusureg, " &_
				"B.apematerno as Apematusureg, " &_
				"C.nombres as Nombreusumod, " &_
				"C.apepaterno as Apepatusumod, " &_
				"C.apematerno as Apematusumod " &_
				"FROM usuario A " &_
				"INNER JOIN usuario B ON B.codusuario=A.usuarioregistra " &_
				"LEFT OUTER JOIN usuario C ON C.codusuario=A.usuariomodifica " &_
				"INNER JOIN UsuarioPerfil D ON D.CodUsuario=A.CodUsuario " &_
				"INNER JOIN Perfil E ON E.CodPerfil=D.CodPerfil " &_
				"WHERE a.codusuario = " & codusuario
				consultar sql,RS
				usuario=rs.Fields("usuario")
				clave=rs.Fields("clave")		
				apepat=rs.Fields("apepaterno")
				apemat=rs.Fields("apematerno")
				nombres=rs.Fields("nombres")	
				correo=rs.Fields("correo")	
				fbloq=rs.Fields("flagbloqueo")
				administrador=rs.Fields("administrador")
				activo=rs.Fields("activo")		
				fechaReg=RS.Fields("fecharegistra")
				usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
				fechaMod=RS.Fields("fechamodifica")
				usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
				codtipousuario=rs.Fields("codtipousuario")			
				RS.Close
			else
				activo="1"					
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codusuario="" then%>Nuevo <%end if%>Usuario</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language="javascript">
				var limpioclave=0;
				<% if codusuario="" then %>
				function agregar()
				{
					if(trim(formula.usuario.value)==""){alert("Debe ingresar el Usuario.");return;}
					if(trim(formula.clave.value)==""){alert("Debe ingresar la Contraseña.");return;}
					if(trim(formula.apepat.value)==""){alert("Debe ingresar el Apellido Paterno del Usuario.");return;}
					if(trim(formula.apemat.value)==""){alert("Debe ingresar el Apellido Materno del Usuario.");return;}
					if(trim(formula.nombres.value)==""){alert("Debe ingresar los Nombres del Usuario.");return;}
					if(!isEmailAddress(formula.correo)){alert("Debe ingresar un e-Mail válido.");return;}
					//agencia
					if(formula.codtipousuario.value=="1" && formula.codagencia.value==""){alert("Debe seleccionar una Agencia de Cobranza.");return;}
					//oficina
					if(formula.codtipousuario.value=="2" && formula.codoficina.value==""){alert("Debe seleccionar una Oficina.");return;}
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<% else %>
				function modificar()
				{
					if(trim(formula.usuario.value)==""){alert("Debe ingresar el Usuario.");return;}
					if(trim(formula.clave.value)==""){alert("Debe ingresar la Contraseña.");return;}
					if(trim(formula.apepat.value)==""){alert("Debe ingresar el Apellido Paterno del Usuario.");return;}
					if(trim(formula.apemat.value)==""){alert("Debe ingresar el Apellido Materno del Usuario.");return;}
					if(trim(formula.nombres.value)==""){alert("Debe ingresar los Nombres del Usuario.");return;}
					if(!isEmailAddress(formula.correo)){alert("Debe ingresar un e-Mail válido.");return;}
					//agencia
					if(formula.codtipousuario.value=="1" && formula.codagencia.value==""){alert("Debe seleccionar una Agencia de Cobranza.");return;}
					//oficina o gestor
					//if((formula.codtipousuario.value=="2"||formula.codtipousuario.value=="3") && formula.codoficina.value==""){alert("Debe seleccionar una Oficina.");return;}
					//el gestor puede ser multioficina o tener oficina seleccionada
					if((formula.codtipousuario.value=="2") && formula.codoficina.value==""){alert("Debe seleccionar una Oficina.");return;}
										
					document.formula.agregardato.value=2;
					document.formula.submit();
				}				
				<% end if %>
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
			<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
					<iframe id="ventanagrab" name="ventanagrab" src="" style="visibility:hidden" width="0" height="0" border="0" frameborder="0"></iframe>
					<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
						<form name="formula" method="post" action="dcs_nuevousuario.asp" target="ventanagrab">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b ><b>&nbsp;<b><%if codusuario="" then%>Nuevo <%end if%>Usuario</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr height=20>
						<td colspan=2 align="right"><font  size=1 color=#483d8b>Registr&oacute;:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modific&aacute;:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<tr>
						<td width=30%><font  size=2 color=#483d8b>&nbsp;&nbsp;Usuario:</font></td>
						<td><input name="usuario" type="text" maxlength=200 value="<%=usuario%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5"><font  size=2 color=#483d8b>&nbsp;&nbsp;Contrase?a:</font></td>
						<td bgcolor="#f5f5f5"><input name="clave" type=password maxlength=200 value="<%=clave%>" style="font-size: xx-small; width: 200px;"  onfocus="if(limpioclave==0){this.value='';limpioclave=1;}"></td>
					</tr>
					<tr>
						<td><font  size=2 color=#483d8b>&nbsp;&nbsp;Apellidos Paterno:</font></td>
						<td><input name="apepat" type="text" maxlength=200 value="<%=apepat%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5"><font  size=2 color=#483d8b>&nbsp;&nbsp;Apellido Materno:</font></td>
						<td bgcolor="#f5f5f5"><input name="apemat" type="text" maxlength=200 value="<%=apemat%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td><font  size=2 color=#483d8b>&nbsp;&nbsp;Nombres:</font></td>
						<td><input name="nombres" type="text" maxlength=200 value="<%=nombres%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5"><font  size=2 color=#483d8b>&nbsp;&nbsp;e-Mail:</font></td>
						<td bgcolor="#f5f5f5"><input name="correo" type="text" maxlength=200 value="<%=correo%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
					<td><font  size=2 color=#483d8b>&nbsp;&nbsp;Tipo Usuario:</font></td>
					<td>
						<select name="codtipousuario" style="font-size: xx-small; width: 200px;" onchange="activarseleccion();">
						<%
						if codusuario<>"" then
							''esto es para que si anteriormente hab?a un tipo de usuario inactivo se muestre
							sql = "SELECT p.CodPerfil AS codtipousuario, p.descripcion "&_
								"FROM UsuarioPerfil a "&_
								"INNER JOIN Perfil p ON p.CodPerfil=a.CodPerfil "&_
								"WHERE activo=1 and p.CodPerfil=" & codtipousuario &_ 
								" ORDER BY descripcion"
							else
							 
							sql = "SELECT p.CodPerfil AS codtipousuario, p.descripcion "&_
								  "FROM UsuarioPerfil a "&_
							      "INNER JOIN Perfil p ON p.CodPerfil=a.CodPerfil "&_
							      "WHERE activo=1 "&_
								  "ORDER BY descripcion"
							"
						end if
						consultar sql,RS
						Do While not RS.EOF
						%>
						<option value="<%=RS.Fields("codtipousuario")%>"<%if codtipousuario<>"" then%><%if RS.fields("codtipousuario")=int(codtipousuario) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>
						</select>
						</td>
					</tr>					
					<tr>
					<td bgcolor="#f5f5f5"><font  size=2 color=#483d8b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Agencia:</font></td>
					<td bgcolor="#f5f5f5">
						<select name = "codagencia" style="font-size: xx-small; width: 200px;">
						<option value="">Seleccionar</option>
						<%
						sql = "select codagencia, razonsocial from agencia order by razonsocial"
						consultar sql,RS
						Do While Not RS.EOF
						%>
						<option value="<%=RS.Fields("codagencia")%>"<% if codagencia<>"" then%><% if RS.fields("codagencia")=int(codagencia) then%> selected<%end if%><%end if%>><%=RS.Fields("razonsocial")%></option>
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>
						</select>
						</td>
					</tr>
					<tr>
					<td bgcolor="#f5f5f5"><font size=2 color=#483d8b>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Oficina:</font></td>
					<td bgcolor="#f5f5f5">
						<select name="codoficina" style="font-size: xx-small; width: 200px;">
						<option value="">Seleccionar</option>
						<%
						sql = "select codoficina, descripcion from oficina order by codoficina"
						consultar sql,RS
						Do While Not RS.EOF
						%>
						<option value="<%=RS.Fields("codoficina")%>"<% if codoficina<>"" then%><% if RS.fields("codoficina")=codoficina then%> selected<%end if%><%end if%>><%=RS.fields("codoficina") & " - " & RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>
						</select>
						</td>
					</tr>					
					<script language="javascript">
					function activarseleccion()
					{
						//agencia
						if(formula.codtipousuario.value=="1")
						{
							formula.codoficina.value="";
							formula.codoficina.disabled=true;
							formula.codagencia.disabled=false;
						}
						//oficina o gestor
						if(formula.codtipousuario.value=="2"||formula.codtipousuario.value=="3")
						{
							formula.codagencia.value="";
							formula.codagencia.disabled=true;
							formula.codoficina.disabled=false;
						}	
						//admin
						if(formula.codtipousuario.value=="4")
						{
							formula.codagencia.value="";
							formula.codoficina.value="";
							formula.codagencia.disabled=true;
							formula.codoficina.disabled=true;
						}																
					}
					activarseleccion();
					</script>
					<tr>
						<td><font  size=2 color=#483d8b>&nbsp;</font></td>
						<td><input type=checkbox name="activo" style="font-size: xx-small;" <%if activo=1 then%> checked<%end if%>>&nbsp;&nbsp;<font  size=2 color=#483d8b>Activo</font></td>
					</tr>					
					<tr>
						<td bgcolor="#f5f5f5"><font  size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#f5f5f5"><input type=checkbox name="fbloq" style="font-size: xx-small;" <%if fbloq=3 then%> checked<%end if%>>&nbsp;&nbsp;<font  size=2 color=#483d8b>Bloqueado</font></td>
					</tr>					
					<tr> 
						<td><font  size="2" color="#483d8b">&nbsp;</font></td>
						<td><input type="checkbox" name="administrador" onclick="document.formula.codperf1.checked=this.checked;" style="font-size: xx-small;" <%if administrador=1 then%> checked<%end if%>>&nbsp;&nbsp;<font  size=2 color=#483d8b>Administrador</font></td>
					</tr>					
					<tr>	
						<td bgcolor="#F5F5F5" colspan="2"><font size="2" color="#483d8b">&nbsp;&nbsp;Perfiles:</font></td>
					</tr>															
							<tr>	
						<td colspan="2">
							<table width="90%" align="center" cellpadding="0" cellspacing="0" border="0">
									<%
									if codusuario<>"" then
										sql="select A.codperfil, CASE WHEN B.activo IS NULL THEN 0 ELSE B.ACTIVO END as activo, A.descripcion, codusuario from perfil A left outer join usuarioperfil B on A.codperfil = B.codperfil and B.codusuario =" & codusuario & " and B.activo = 1 order by A.codperfil"
									else
										sql="select codperfil, descripcion, 0 as activo, null as codusuario from perfil order by codperfil"
									end if
									consultar sql,RS
										perfil=""
										cadenaperfil=""
										contadorperfil=0
										intercalacolor=""
										Do While Not RS.EOF
												perfil=RS.Fields("descripcion")
												if RS.Fields("codperfil")=1 then
												cadenaperfil=cadenaperfil & "<tr " & intercalacolor & "><td valign=top><font  size=1 color=#483d8b>&nbsp;&nbsp;</font></td><td><input type=checkbox onclick='document.formula.administrador.checked=this.checked;' name='codperf" & RS.Fields("codperfil") & "' style='font-size: xx-small;' " & iif(RS.Fields("activo")=1,"checked","") & "><font  size=1 color=#483d8b>" & RS.Fields("descripcion") & "</font></td></tr>" & chr(10)
												else
												cadenaperfil=cadenaperfil & "<tr " & intercalacolor & "><td valign=top><font  size=1 color=#483d8b>&nbsp;&nbsp;</font></td><td><input type=checkbox name='codperf" & RS.Fields("codperfil") & "' style='font-size: xx-small;' " & iif(RS.Fields("activo")=1,"checked","") & "><font  size=1 color=#483d8b>" & RS.Fields("descripcion") & "</font></td></tr>" & chr(10)
												end if
												contadorperfil=contadorperfil + 1	
												if intercalacolor="" then
													intercalacolor=" bgcolor='#F5F5F5' "
												else
													intercalacolor=""
												end if																	
										RS.MoveNext
										Loop
										RS.Close	
									%>	
									<%=cadenaperfil%>
							</table>				
						</td>
					</tr>	
					<tr>					
						<td bgcolor="#F5F5F5"><font size="2" color="#483d8b">&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align="right" height="40"><%if codusuario="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
					
							<input type="hidden" name="agregardato" value="">
							<input type="hidden" name="hclave" value="<%=clave%>">
							<input type="hidden" name="codusuario" value="<%=codusuario%>">
							<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
							<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
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
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>

