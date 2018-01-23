<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admfamproducto.asp") then
	buscador=obtener("buscador")	
	codfamproducto=obtener("codfamproducto")
		if obtener("agregardato")<>"" then
		codproducto=obtener("codproducto")
		descripcion=obtener("descripcion")
		if obtener("activo")<>"" then activo="1" else activo="0" end if	
													
			existefamproducto=0
			
			if codfamproducto<>"" then
				sql="select count(*) from famproducto where descripcion='" & descripcion & "' and codfamproducto<>" & codfamproducto & " and codproducto='" & codproducto & "'"
			else			
				sql="select count(*) from famproducto where descripcion='" & descripcion & "' and codproducto='" & codproducto & "'"
			end if
			consultar sql,RS
			existefamproducto=RS.Fields(0)
			RS.Close			
			if existefamproducto=0 then			
				if obtener("agregardato")="1" then		
					
						sql="insert into famproducto (codproducto,descripcion,activo,usuarioregistra,fecharegistra) values ('" & codproducto & "','" & descripcion & "'," & activo & "," & session("codusuario") & ",getdate())"
				
				else
					
						sql="update famproducto set codproducto='" & codproducto & "',descripcion='" & descripcion & "',activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codfamproducto=" & codfamproducto
					
				end if
			conn.execute sql
				''Response.Write sql
									
				%>
				<script language=javascript>
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
			
					<%end if%>				
					<%if obtener("paginapadre")="admfamproducto.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
				    alert("La familia ya existe.");
				    history.back();
				</script>			
			<%				
			end if
		else
			if codfamproducto<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from famproducto A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where A.codfamproducto=" & codfamproducto 
					consultar sql,RS
					descripcion=rs.Fields("descripcion")
					codproducto=rs.Fields("codproducto")		
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
			<title><%if codfamproducto="" then%>Nueva <%end if%>Familia Producto</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				<%if codfamproducto="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción para ingreso al Sistema.");return;}
																	
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción para ingreso al Sistema.");return;}
					
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
					<form name=formula method=post action="nuevofamproducto.asp">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codfamproducto="" then%>Nueva <%end if%>Familia Producto</b></b></font>
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
					<td bgcolor ="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Producto:</font></td>
					<td bgcolor ="#f5f5f5">
						<select name="codproducto" style="font-size: xx-small; width: 200px;">
						<%
						sql = "select codproducto,descripcion from producto order by codproducto"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codproducto")%>" <%if codproducto<>"" then%><% if RS.fields("codproducto")=codproducto then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>
						</select>
						</td>
					</tr>
					
					
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descripción:</font></td>
						<td> <input name="descripcion" type=text maxlength=50 value="<%=descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td width=30% bgcolor ="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Estado:</font></td>
						<td bgcolor ="#f5f5f5"><input type=checkbox name="activo" style="font-size: xx-small;" <%if activo=1 then%> checked<%end if%>>&nbsp;&nbsp;<font face=Arial size=2 color=#483d8b>Activo</font></td>
					</tr>			
					<tr>					
						<td><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td align=right height=40><%if codfamproducto="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
						<input type=hidden name="agregardato" value="">
						<input type=hidden name="vistapadre" value="<%=obtener("vistapadre")%>">
						<input type=hidden name="paginapadre" value="<%=obtener("paginapadre")%>">
						<input type=hidden name="codfamproducto" value="<%=codfamproducto%>">
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
	    window.open("userexpira.asp", "_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
    alert("Tiempo Expirado");
    window.open("index.html","_top");
    window.close();
</script>
<%
end if
%>

