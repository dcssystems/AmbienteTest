<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admparametro.asp") then
	buscador=obtener("buscador")	
	codparametro=obtener("codparametro")
		if obtener("agregardato")<>"" then
		descripcion=obtener("descripcion")
		valortexto1=obtener("valortexto1")
		valortexto2=obtener("valortexto2")
		valortexto3=obtener("valortexto3")
		valortexto4=obtener("valortexto4")
		valornumerico1=obtener("valornumerico1")
		valornumerico2=obtener("valornumerico2")
		valornumerico3=obtener("valornumerico3")
		valornumerico4=obtener("valornumerico4")
		
		if valornumerico1 = "" then
		xvalornumerico1 = "null"
		else
		xvalornumerico1 = cdbl(valornumerico1)
		end if	
		if valornumerico2 = "" then
		xvalornumerico2 = "null"
		else
		xvalornumerico2 = cdbl(valornumerico2)
		end if	
		if valornumerico3 = "" then
		xvalornumerico3 = "null"
		else
		xvalornumerico3 = cdbl(valornumerico3)
		end if	
		if valornumerico4 = "" then
		xvalornumerico4 = "null"
		else
		xvalornumerico4 = cdbl(valornumerico4)
		end if	
									
			existeparametro=0
			
			if codparametro<>"" then
			sql="select count(*) from parametro where descripcion='" & descripcion & "' and codparametro<>" & codparametro 
			else
			sql="select count(*) from parametro where descripcion='" & descripcion & "'"
			end if
			consultar sql,RS
			existeparametro=RS.Fields(0)
			RS.Close			
			if existeparametro=0 then			
				if obtener("agregardato")="1" then		
					sql="insert into parametro (descripcion,valortexto1,valortexto2,valortexto3,valortexto4,valornumerico1,valornumerico2,valornumerico3,valornumerico4,usuarioregistra,fecharegistra) values ('" & descripcion & "','" & valortexto1 & "','" & valortexto2 & "','" & valortexto3 & "','" & valortexto4 & "'," & xvalornumerico1 & "," & xvalornumerico2 & "," & xvalornumerico3 & "," & xvalornumerico4 & "," & session("codusuario") & ",getdate())"
				else
					sql="update parametro set descripcion='" & descripcion & "', valortexto1 = '"& valortexto1 &"' , valortexto2 = '"& valortexto2 &"', valortexto3 = '"& valortexto3 &"', valortexto4 = '"& valortexto4 &"', valornumerico1 = " & xvalornumerico1 & " , valornumerico2 = " & xvalornumerico2 & " , valornumerico3 = " & xvalornumerico3 & " , valornumerico4 = " & xvalornumerico4 & ",    usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codparametro=" & codparametro
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
					<%if obtener("paginapadre")="admparametro.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
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
			if codparametro<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from parametro A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where a.codparametro = " & codparametro
					consultar sql,RS
					descripcion=rs.Fields("descripcion")	
					valortexto1=rs.Fields("valortexto1")
					valortexto2=rs.Fields("valortexto2")
					valortexto3=rs.Fields("valortexto3")
					valortexto4=rs.Fields("valortexto4")
					valornumerico1=rs.Fields("valornumerico1")
					valornumerico2=rs.Fields("valornumerico2")
					valornumerico3=rs.Fields("valornumerico3")
					valornumerico4=rs.Fields("valornumerico4")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codparametro="" then%>Nuevo <%end if%>Grupo Territorio</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
				var limpioclave=0;
				<%if codparametro="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(isNaN(trim(formula.valornumerico1.value.replace(",","")))){alert("El valornumerico1 debe ser un dato numérico.");return;}
					if(isNaN(trim(formula.valornumerico2.value.replace(",","")))){alert("El valornumerico2 debe ser un dato numérico.");return;}
					if(isNaN(trim(formula.valornumerico3.value.replace(",","")))){alert("El valornumerico3 debe ser un dato numérico.");return;}
					if(isNaN(trim(formula.valornumerico4.value.replace(",","")))){alert("El valornumerico4 debe ser un dato numérico.");return;}
																							
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(isNaN(trim(formula.valornumerico1.value.replace(",","")))){alert("El valornumerico1 debe ser un dato numérico.");return;}
					if(isNaN(trim(formula.valornumerico2.value.replace(",","")))){alert("El valornumerico2 debe ser un dato numérico.");return;}
					if(isNaN(trim(formula.valornumerico3.value.replace(",","")))){alert("El valornumerico3 debe ser un dato numérico.");return;}
					if(isNaN(trim(formula.valornumerico4.value.replace(",","")))){alert("El valornumerico4 debe ser un dato numérico.");return;}
																			
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
					<form name=formula method=post action="nuevoparametro.asp">
					<tr height=20>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codparametro="" then%>Nuevo <%end if%>Parámetro</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr height=20>
						<td colspan=2 align=right><font face=Arial size=1 color=#483d8b>Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<!--<tr height=20>
						<td><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descripción:</font></td>
						<td><input name="descripcion" type=text maxlength=50 value="<%=descripcion%>" style="font-size: xx-small; width: 250px;"></td>
					</tr>-->
					<tr height=20>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Descripción:</font></td>
						<td><input name="descripcion" type=text maxlength=200 value="<%=descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr height=20>
						<td bgcolor="#F5F5F5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valortexto1:</font></td>
						<td bgcolor="#F5F5F5"><input name="valortexto1" type=text maxlength=1000 value="<%=valortexto1%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr height=20>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valortexto2:</font></td>
						<td><input name="valortexto2" type=text maxlength=1000 value="<%=valortexto2%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr height=20>
						<td  bgcolor="#F5F5F5"width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valortexto3:</font></td>
						<td bgcolor="#F5F5F5"><input name="valortexto3" type=text maxlength=1000 value="<%=valortexto3%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr height=20>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valortexto4:</font></td>
						<td><input name="valortexto4" type=text maxlength=1000 value="<%=valortexto4%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr height=20>
						<td bgcolor="#F5F5F5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valornumerico1:</font></td>
						<td bgcolor="#F5F5F5"><input name="valornumerico1" type=text maxlength=1000 value="<%=valornumerico1%>" style="font-size: xx-small; width: 75px;"></td>
					</tr>
					<tr height=20>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valornumerico2:</font></td>
						<td><input name="valornumerico2" type=text maxlength=1000 value="<%=valornumerico2%>" style="font-size: xx-small; width: 75px;"></td>
					</tr>
					<tr height=20>
						<td bgcolor="#F5F5F5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valornumerico3:</font></td>
						<td bgcolor="#F5F5F5"><input name="valornumerico3" type=text maxlength=1000 value="<%=valornumerico3%>" style="font-size: xx-small; width: 75px;"></td>
					</tr>
					<tr height=20>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;valornumerico4:</font></td>
						<td><input name="valornumerico4" type=text maxlength=1000 value="<%=valornumerico4%>" style="font-size: xx-small; width: 75px;"></td>
					</tr>
					<tr>					
						<td bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td bgcolor="#F5F5F5" align=right height=40><%if codparametro="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
						<input type=hidden name="agregardato" value="">
						<input type=hidden name="codparametro" value="<%=codparametro%>">
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

