<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admoficina.asp") then
	buscador=obtener("buscador")	
	codoficina=obtener("codoficina")
		if obtener("agregardato")<>"" then
			codterritorio=obtener("codterritorio")
			descripcion=obtener("descripcion")
			if obtener("activo")<>"" then activo=1 else activo=0 end if	
			codgrupoubigeo = obtener("codgrupoubigeo")
			coddpto = obtener("coddpto")
			if obtener("codprov")<>"" then
			codprov = mid(obtener("codprov"),3,2)
			end if
			if obtener("coddist")<>"" then
			coddist = mid(obtener("coddist"),5,3)
			end if
			codoficinanuevo=obtener("codoficinanuevo")
			
			existeoficina=0
			
			if codoficinanuevo<>"" then
			sql="select count(*) from oficina where codoficina = '" & codoficina & "' or descripcion='" & descripcion & "'"
			else
			sql="select count(*) from oficina where descripcion='" & descripcion & "' and codoficina <> '" & codoficina & "'"
			end if
			consultar sql,RS
			existeoficina=RS.Fields(0)
			RS.Close			
			if existeoficina=0 then			
			    if coddpto<>"" then
			        xcoddpto="'" & coddpto & "'"
			    else
			        xcoddpto="null"
			    end if
                if codprov<>"" then
			        xcodprov="'" & codprov & "'"
			    else
			        xcodprov="null"
			    end if	
			    if coddist<>"" then
			        xcoddist="'" & coddist & "'"
			    else
			        xcoddist="null"
			    end if			    
			        
				if obtener("agregardato")="1" then		
				sql="insert into oficina (codoficina,codterritorio,descripcion,activo,coddpto,codprov,coddist,usuarioregistra,fecharegistra) values ('" & codoficina & "','" & codterritorio & "','" & descripcion & "'," & activo & "," & xcoddpto & "," & xcodprov & "," & xcoddist & "," & session("codusuario") & ",getdate())"
				else
				sql="update oficina set codterritorio = '" & codterritorio & "',descripcion='" & descripcion & "',activo=" & activo & ",coddpto =" & xcoddpto & ",codprov =" & xcodprov & ",coddist =" & xcoddist & ", usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codoficina='" & codoficina & "'"
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
					<%if obtener("paginapadre")= "admoficina.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("La Oficina ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codoficina<>"" then
					sql="select A.*,D.CodGrupoUbigeo,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from oficina A inner join usuario B on B.codusuario=A.usuarioregistra left outer join Ubigeo D on A.coddpto = D.coddpto and A.codprov = D.codprov and A.coddist = D.coddist left outer join usuario C on C.codusuario=A.usuariomodifica where A.codoficina= '" & codoficina & "'"
					consultar sql,RS
					if obtener("ubigeoact")<>"" then
						descripcion=obtener("descripcion")
						codterritorio=obtener("codterritorio")							
						if obtener("activo")<>"" then activo=1 else activo=0 end if	
						codgrupoubigeo = obtener("codgrupoubigeo")
						coddpto = obtener("coddpto")
						codprov = obtener("codprov")
						coddist = obtener("coddist")						
					else
						descripcion=rs.Fields("descripcion")
						codterritorio=rs.Fields("codterritorio")							
						activo=rs.Fields("activo")										
						codgrupoubigeo = RS.fields("codgrupoubigeo")
						coddpto = rs.Fields("coddpto")
						codprov = rs.Fields("coddpto") & rs.Fields("codprov")
						coddist = rs.Fields("coddpto") & rs.Fields("codprov") & rs.Fields("coddist")											
					end if
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			else
			        codoficinanuevo=obtener("codoficinanuevo")
			        codterritorio=obtener("codterritorio")
			        descripcion=obtener("descripcion")
			        if obtener("ubigeoact")<>"" then
			            if obtener("activo")<>"" then activo=1 else activo=0 end if
			        else
			        activo=1
			        end if
					codgrupoubigeo = obtener("codgrupoubigeo")
					coddpto = obtener("coddpto")
					codprov = obtener("codprov")
					coddist = obtener("coddist")	
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codoficina="" then%>Nueva <%end if%>Oficina</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				<%if codoficina="" then%>
				function agregar()
				{
					document.formula.codoficina.value=document.formula.codoficinanuevo.value;				
					if(trim(formula.codoficina.value)==""){alert("Debe ingresar un Código para el ingreso al Sistema.");return;}
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción para ingreso al Sistema.");return;}
				
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.codoficina.value)==""){alert("Debe ingresar un Código para ingreso al Sistema.");return;}
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción para ingreso al Sistema.");return;}
					
					document.formula.agregardato.value=2;
					document.formula.submit();
				}				
				<%end if%>
				function actualizaubigeo()
				{
				document.formula.ubigeoact.value = 1;
				document.formula.agregardato.value=""
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
				<form name=formula action="nuevooficina.asp"><!--method=post : sólo para este caso porque no funca el history.back por post en extranet-->
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codoficina="" then%>Nueva <%end if%>Oficina</b></b></font>
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
						<td><%if codoficina = "" then%><input name="codoficinanuevo" type=text maxlength=50 value="<%=codoficinanuevo%>" style="font-size: xx-small; width: 100px;"></td>
							  <%else%><font face=Arial size=2 color=#483d8b><b><%=codoficina%></b></font></td>
							  <%end if%>
				    </tr>
					<tr>
						<td width=30% bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descripcion:</font></td>
						<td bgcolor ="#f5f5f5"><input name="descripcion" type=text maxlength=200 value="<%=descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Territorio:</font></td>
					<td>
						<select name="codterritorio" style="font-size: xx-small; width: 200px;">
						<%
						sql = "select codterritorio,descripcion from Territorio order by descripcion"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codterritorio")%>" <% if RS.fields("codterritorio")=codterritorio then%> selected<%end if%>><%=RS.fields("codterritorio") & " - " & RS.Fields("descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>
						</select>
						</td>
					</tr>												  
					<tr>
					<td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Grupo Ubigeo:</font></td>
					<td bgcolor="#f5f5f5">
						<select name="codgrupoubigeo" onchange="actualizaubigeo()" style="font-size: xx-small; width: 200px;">
						<%
						sql = "select codgrupoubigeo,descripcion from GrupoUbigeo order by codgrupoubigeo"
						consultar sql,RS
                        seleccion=0
						primercodgrupoubigeo=""						
						Do While Not  RS.EOF
							if primercodgrupoubigeo="" then
								primercodgrupoubigeo=RS.Fields("codgrupoubigeo")
							end if						
						%>
						<option value="<%=RS.Fields("codgrupoubigeo")%>" <%if RS.fields("codgrupoubigeo")=codgrupoubigeo then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						if seleccion=0 then
							codgrupoubigeo=primercodgrupoubigeo
						end if						
						%>
						</select>
						</td>
					</tr>
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Departamento:</font></td>
					<td>
						<select name="coddpto" onchange="actualizaubigeo()" style="font-size: xx-small; width: 200px;">
						<option value="">Seleccionar</option>
						<%
						sql = "select coddpto,departamento from ubigeo where codgrupoubigeo = '" & codgrupoubigeo & "' and departamento<>'' and coddpto<>'00' group by coddpto,departamento order by departamento"
						consultar sql,RS
						seleccion=0
						primercoddpto=""
						Do While Not  RS.EOF
							if primercoddpto="" then
								primercoddpto=RS.Fields("coddpto")
							end if
						%>
						<option value="<%=RS.Fields("coddpto")%>" <%if RS.fields("coddpto")=coddpto then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("departamento")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						if seleccion=0 then
							''coddpto=primercoddpto
						end if
						%>
						</select>
						</td>
					</tr>					
					<tr>
					<td bgcolor ="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Provincia:</font></td>
					<td bgcolor ="#f5f5f5">
						<select name = "codprov" onchange="actualizaubigeo()" style="font-size: xx-small; width: 200px;">
						<option value="">Seleccionar</option>
						<%
						if coddpto<>"" then
						    sql = "select codprov,provincia from ubigeo where codgrupoubigeo = '" & codgrupoubigeo & "' and coddpto = '" & coddpto & "' and provincia<>'' and codprov<>'00' group by codprov,provincia order by provincia"
						    consultar sql,RS
						    seleccion=0
						    primercodprov=""
						    Do While Not RS.EOF
							    if primercodprov="" then
								    primercodprov=coddpto & RS.Fields("codprov")
							    end if
						    %>
						    <option value="<%=coddpto & RS.Fields("codprov")%>" <%if coddpto & RS.fields("codprov")=codprov then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("provincia")%></option>
						    <%
						    RS.MoveNext
						    loop
						    RS.Close
						    if seleccion=0 then
							    ''codprov=primercodprov
						    end if						
						end if
						%>
						</select>
						</td>
					</tr>
					<tr>
					<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Distrito:</font></td>
					<td>
						<select name = "coddist" onchange="actualizaubigeo()" style="font-size: xx-small; width: 200px;">
						<option value="">Seleccionar</option>
						<%
						if  codprov<>"" then
						    sql = "select coddist,distrito from ubigeo where codgrupoubigeo = '" & codgrupoubigeo & "' and coddpto = '" & coddpto & "' and coddpto + codprov = '" & codprov & "' and distrito<>'' and coddist<>'000' group by coddist,distrito order by distrito"
						    consultar sql,RS
						    seleccion=0
						    primercoddist=""
						    Do While Not RS.EOF
							    if primercoddist="" then
								    primercoddist=codprov & RS.Fields("coddist")
							    end if
						    %>
						    <option value="<%=codprov & RS.Fields("coddist")%>" <%if codprov & RS.fields("coddist")=coddist then%> selected<%seleccion=1%><%end if%>><%=RS.Fields("distrito")%></option>
						    <%
						    RS.MoveNext
						    loop
						    RS.Close
						    if seleccion=0 then
							    ''coddist=primercoddist
						    end if	
						end if						
						%>
						</select>
						</td>
					</tr>
					<tr>
						<td bgcolor="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Activo:</font></td>
						<td bgcolor="#f5f5f5"><input type=checkbox name="activo" style="font-size: xx-small;" <%if activo=1 then%> checked<%end if%>>&nbsp;&nbsp;<font face=Arial size=2 color=#483d8b>Activo</font></td>
					</tr>			
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codoficina="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
					</table>
					<input type=hidden name="agregardato" value="">
					<input type=hidden name="ubigeoact" value="">
					<input type=hidden name="codoficina" value="<%=codoficina%>">
					<input type=hidden name="vistapadre" value="<%=obtener("vistapadre")%>">
					<input type=hidden name="paginapadre" value="<%=obtener("paginapadre")%>">
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

