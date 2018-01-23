<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admperfil.asp") then	
	codperfil=obtener("codperfil")
		if obtener("agregardato")<>"" then
			descripcion=obtener("descripcion")
			orden = obtener("orden")
			
			existeperfil=0
			
			if codperfil<>"" then
			sql="select count(*) from Perfil where descripcion='" & descripcion & "' and codperfil<>" & codperfil
			else
			sql="select count(*) from Perfil where descripcion='" & descripcion & "'"
			end if
			consultar sql,RS
			existeperfil=RS.Fields(0)
			RS.Close			
			
			if existeperfil=0 then			
				if obtener("agregardato")="1" then		
				sql="insert into perfil (descripcion,usuarioregistra,fecharegistra,orden) values ('" & descripcion & "'," & session("codusuario") & ",getdate()," & orden & ")"
				else
				sql="update perfil set descripcion='" & descripcion & "', orden='"& orden &"' ,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codperfil=" & codperfil
				end if
				''Response.Write sql
				conn.execute sql
				if codperfil="" then
					sql="select codperfil from Perfil where descripcion='" & descripcion & "'"
					consultar sql,RS
					codperfil=RS.Fields(0)
					RS.Close	
				end if
				
				''aquí se elimina y se agregan  las facultades seleccionadas
				sql="delete perfilfacultad where codperfil=" & codperfil
				conn.execute sql
				sql="select A.CodGrupoFacultad,A.CodFacultad,A.Descripcion as Facultad,B.Descripcion as GrupoFacultad,null as codperfil from Facultad A inner join GrupoFacultad B on A.codgrupofacultad=B.codgrupofacultad order by A.CodGrupoFacultad,A.CodFacultad"
				consultar sql,RS
				Do While not RS.EOF
					if obtener("codfact" & RS.Fields("CodFacultad"))<>"" then
						sql="Insert into PerfilFacultad (codperfil,codfacultad) values (" & codperfil & "," & RS.Fields("CodFacultad") & ")"
						conn.execute sql
					end if
				RS.MoveNext
				Loop
				RS.Close
				%>
				<script language=javascript>
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="admperfil.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("El perfil ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codperfil<>"" then
				sql="select A.descripcion,A.orden,A.fecharegistra,B.nombres as nombresReg,B.apepaterno as apepaternoReg,B.apematerno as apematernoReg,A.fechamodifica,C.nombres as nombresMod,C.apepaterno as apepaternoMod,C.apematerno as apematernoMod from perfil A left outer join Usuario B on A.usuarioregistra=B.codusuario left outer join Usuario C on A.usuariomodifica=C.codusuario where A.codperfil=" & codperfil
				consultar sql,RS
					descripcion=rs.Fields("descripcion")	
					orden = rs.fields("orden")	
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("nombresReg")),"",RS.Fields("nombresReg")) & ", " & iif(IsNull(RS.Fields("apepaternoReg")),"",RS.Fields("apepaternoReg")) & " " & iif(IsNull(RS.Fields("apematernoReg")),"",RS.Fields("apematernoReg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("nombresMod")),"",RS.Fields("nombresMod")) & ", " & iif(IsNull(RS.Fields("apepaternoMod")),"",RS.Fields("apepaternoMod")) & " " & iif(IsNull(RS.Fields("apematernoMod")),"",RS.Fields("apematernoMod"))
					''apepat=rs.Fields("apepaterno")		
					''apemat=rs.Fields("apematerno")
					''nombres=rs.Fields("nombres")		
					''correo=rs.Fields("correo")		
					''fbloq=rs.Fields("flagbloqueo")	
					''administrador=rs.Fields("administrador")	
					''activo=rs.Fields("activo")		
					RS.Close
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codperfil="" then%>Nuevo <%end if%>Perfil</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				<%if codperfil="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
					//if(trim(formula.clave.value)==""){alert("Debe ingresar el Password para ingreso al Sistema.");return;}
					//if(trim(formula.apepat.value)==""){alert("Debe ingresar el Apellido Paterno del Usuario a Ingresar.");return;}
					//if(trim(formula.apemat.value)==""){alert("Debe ingresar el Apellido Materno del Usuario a Ingresar.");return;}
					//if(trim(formula.nombres.value)==""){alert("Debe ingresar los Nombres del Usuario a Ingresar.");return;}
					//if(!isEmailAddress(formula.correo)){alert("Debe ingresar un e-mail del Usuario a Ingresar.");return;}
															
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
					//if(trim(formula.clave.value)==""){alert("Debe ingresar el Password para ingreso al Sistema.");return;}
					//if(trim(formula.apepat.value)==""){alert("Debe ingresar el Apellido Paterno del Usuario a Ingresar.");return;}
					//if(trim(formula.apemat.value)==""){alert("Debe ingresar el Apellido Materno del Usuario a Ingresar.");return;}
					//if(trim(formula.nombres.value)==""){alert("Debe ingresar los Nombres del Usuario a Ingresar.");return;}
					//if(!isEmailAddress(formula.correo)){alert("Debe ingresar un e-mail del Usuario a Ingresar.");return;}
					
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
					<form name=formula method=post action="nuevoperfil.asp">
					<tr height=20>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codperfil="" then%>Nuevo <%end if%>Perfil</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr height=20>
						<td colspan=2 align=right><font face=Arial size=1 color=#483d8b>Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<tr height=20>
						<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descripción:</font></td>
						<td><input name="descripcion" type=text maxlength=200 value="<%=descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr height=20>
						<td bgcolor="#F5F5F5" ><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Orden:</font></td>
						<td bgcolor="#F5F5F5" ><input name="orden" type=text maxlength=50 value="<%=orden%>" style="font-size: xx-small; width: 60px; text-align: right"></td>
					</tr>		
					<tr>	
						<td colspan=2><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Facultades del Perfil:</font></td>
					</tr>				
					<tr>	
						<td colspan=2>
							<table width=90% align=center cellpadding=0 cellspacing=0 border=0>
									<%
									if codperfil<>"" then
										sql="select A.CodGrupoFacultad,A.CodFacultad,A.Descripcion as Facultad,B.Descripcion as GrupoFacultad,C.codperfil from Facultad A inner join GrupoFacultad B on A.codgrupofacultad=B.codgrupofacultad left outer join PerfilFacultad C on A.codfacultad=C.codfacultad and C.codperfil=" & codperfil & " order by A.CodGrupoFacultad,A.CodFacultad"
									else
										sql="select A.CodGrupoFacultad,A.CodFacultad,A.Descripcion as Facultad,B.Descripcion as GrupoFacultad,null as codperfil from Facultad A inner join GrupoFacultad B on A.codgrupofacultad=B.codgrupofacultad order by A.CodGrupoFacultad,A.CodFacultad"
									end if
									consultar sql,RS
										GrupoFactultad=""
										cadenagrupofacultad=""
										contadorfacultad=0
										intercalacolor=""
										Do While Not RS.EOF
											'if GrupoFactultad="" then
											'	GrupoFactultad=RS.Fields("GrupoFacultad")
											'	cadenagrupofacultad=cadenagrupofacultad & "<tr " & intercalacolor & "><td valign=top rowspan=numfilas><font face=Arial size=1 color=#483d8b>&nbsp;&nbsp;<b>" & GrupoFactultad & "</b></font></td><td><input type=checkbox name='codfact" & RS.Fields("CodFacultad") & "' style='font-size: xx-small;' " & iif(IsNull(RS.Fields("codperfil")),"","checked") & "><font face=Arial size=1 color=#483d8b>" & RS.Fields("Facultad") & "</font></td></tr>" & chr(10)
											'	contadorfacultad=contadorfacultad + 1
											'end if
											if GrupoFactultad<>RS.Fields("GrupoFacultad") then
												GrupoFactultad=RS.Fields("GrupoFacultad")
												cadenagrupofacultad=replace(cadenagrupofacultad,"numfilas",contadorfacultad)
												contadorfacultad=0
												if intercalacolor="" then
													intercalacolor=" bgcolor='#F5F5F5' "
												else
													intercalacolor=""
												end if
												cadenagrupofacultad=cadenagrupofacultad & "<tr " & intercalacolor & "><td valign=top rowspan=numfilas><font face=Arial size=1 color=#483d8b>&nbsp;&nbsp;<b>" & GrupoFactultad & "</b></font></td><td><input type=checkbox name='codfact" & RS.Fields("CodFacultad") & "' style='font-size: xx-small;' " & iif(IsNull(RS.Fields("codperfil")),"","checked") & "><font face=Arial size=1 color=#483d8b>" & RS.Fields("Facultad") & "</font></td></tr>" & chr(10)
												contadorfacultad=contadorfacultad + 1
											else
												cadenagrupofacultad=cadenagrupofacultad & "<tr " & intercalacolor & "><td><input type=checkbox name='codfact" & RS.Fields("CodFacultad") & "' style='font-size: xx-small;' " & iif(IsNull(RS.Fields("codperfil")),"","checked") & "><font face=Arial size=1 color=#483d8b>" & RS.Fields("Facultad") & "</font></td></tr>" & chr(10)
												contadorfacultad=contadorfacultad + 1
											end if											
										RS.MoveNext
										Loop
									RS.Close	
									cadenagrupofacultad=replace(cadenagrupofacultad,"numfilas",contadorfacultad)
									%>	
									<%=cadenagrupofacultad%>
							</table>				
						</td>
					</tr>
					<tr>	
							
					    <td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codperfil="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>	
					<input type=hidden name="agregardato" value="">
					<input type=hidden name="codperfil" value="<%=codperfil%>">
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
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>

