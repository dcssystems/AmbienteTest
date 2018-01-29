﻿<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admCliente.asp") then
		buscador=obtener("buscador")	
		idcliente=obtener("IDCliente")
		if obtener("agregardato")<>"" then
			descripcion=obtener("descripcion")
			orden=obtener("orden")
			if not isNumeric(orden) then
				orden="0"
			end if	
			orden=int(orden)
										
									
			existegrupofacultad=0
			
			if idcliente<>"" then
			sql="select count(*) from grupofacultad where descripcion='" & descripcion & "' and codgrupofacultad<>" & idcliente 
			else
			sql="select count(*) from grupofacultad where descripcion='" & descripcion & "'"
			end if
			consultar sql,RS
			existegrupofacultad=RS.Fields(0)
			RS.Close			
			if existegrupofacultad=0 then			
				if obtener("agregardato")="1" then		
					sql="insert into grupofacultad (descripcion,orden,usuarioregistra,fecharegistra) values ('" & descripcion & "'," & orden & "," & session("codusuario") & ",getdate())"
				else
					sql="update grupofacultad set descripcion='" & descripcion & "',orden=" & orden & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where IDCliente = " & idcliente
				end if
				''Response.Write sql
				conn.execute sql
									
				%>
				<script language="javascript">
					<%if obtener("agregardato")="1" then%>
					//alert("Se agreg&oacute; el usuario correctamente.");
					<%else%>
					//alert("Se modific&oacute; el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="dcs_admCliente.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language="javascript">
					alert("El grupo ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if idcliente<>"" then
					sql="SELECT A.*, " & _
						"	B.nombres AS Nombreusureg, " & _
						"	B.apepaterno AS Apepatusureg, " & _
						"	B.apematerno AS Apematusureg, " & _
						"	C.nombres AS Nombreusumod, " & _
						"	C.apepaterno AS Apepatusumod, " & _
						"	C.apematerno AS Apematusumod " & _
						"FROM cliente A " & _
						"INNER JOIN usuario B ON B.codusuario=A.UsuarioRegistra " & _
						"LEFT OUTER JOIN usuario C ON C.codusuario=A.UsuarioModifica " & _
						"WHERE A.IDCliente = " & idcliente 
					consultar sql,RS
					razonsocial= rs.Fields("RazonSocial")	
					ruc        = rs.Fields("RUC")
					direccion  = rs.Fields("Direccion")	
					telefono   = rs.Fields("Telefono")
					activo     = rs.Fields("Activo")					
					fechaReg   = RS.Fields("FechaRegistra")
					usuarioReg = iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod   = RS.Fields("FechaModifica")
					usuarioMod = iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<head>
			<title><%if idcliente="" then%>Nuevo <%end if%>Cliente</title>
			
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->
			
			<script language='javascript' src="scripts/popcalendar_cobcm.js"></script> 
			<script language="javascript">
				var limpioclave=0;
				<%if idcliente="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripci&oacute;n.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripci&oacute;n.");return;}
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
			</head>
			<body topmargin="0" leftmargin="0">
				<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
					<form name="formula" method="post" action="dcs_nuevoCliente.asp">					
						<tr class="fondo-red">	
							<td class="text-withe" colspan="2">			
								<font size="2" ><b>&nbsp;<b><%if idcliente="" then%>Nuevo <%end if%>Cliente</b></b></font>
							</td>
						</tr>
						<%if fechaReg<>"" then%>
						<tr height="20">
							<td class="text-orange label-registra" colspan="2" align="right"><font size="1">Registr&oacute;:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
							<%if fechaMod<>"" then%><BR>Modific&oacute;:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
							</font></td>
						</tr>	
						<%end if%>	
						<!-- SECCION DEL FORMULARIO PARA CLIENTE -->						
						<tr class="fondo-gris">
							<td class="text-orange-form" width="30%"> Raz&oacute;n Social:</td>
							<td><input name="razonsocial" type="text" maxlength="200" value="<%=razonsocial%>" style="font-size: xx-small; width: 200px;"></td>
						</tr>
						<tr>
							<td class="text-orange" width="30%"><font size="2"> Documento de Identidad: </font></td>
							<td><input name="ruc" type="text" maxlength="200" value="<%=ruc%>" style="font-size: xx-small; width: 200px; text-align: left"></td>
						</tr>
						<tr class="fondo-gris">
							<td class="text-orange" width="30%"><font size="2"> Direcci&oacute;n:</font></td>
							<td><input name="direccion" type="text" maxlength="200" value="<%=direccion%>" style="font-size: xx-small; width: 200px;"></td>
						</tr>
						<tr>
							<td class="text-orange" width="30%"><font size="2"> Tel&eacute;fono:</font></td>
							<td><input name="telefono" type="text" maxlength="200" value="<%=telefono%>" style="font-size: xx-small; width: 200px; text-align: left"></td>
						</tr>
						<tr class="fondo-gris">
							<td class="text-orange" width="30%"><font size="2"> Email:</font></td>
							<td><input name="email" type="text" maxlength="200" value="<%=email%>" style="font-size: xx-small; width: 200px;"></td>
						</tr>
						<tr>
							<td class="text-orange" width="30%"><font size="2"> Activo:</font></td>
							<td><input name="activo" type="checkbox" value="<%=activo%>" <%=iif(IsNull(activo), "", "checked")%>  /></td>
						</tr>				
						
						<!-- FRIN SECCION DEL FORMULARIO PARA CLIENTE -->

					
						<tr class="fondo-red">					
							<td><font size=2 >&nbsp;</font></td>
							<td align=right height=40>
								<%if idcliente="" then%>
								<a href="javascript:agregar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
								<%else%>
								<a href="javascript:modificar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
								<%end if%>
								<a href="javascript:window.close();"><i class="logout demo-icon icon-logout">&#xe800;</i></a>&nbsp;
							</td>					
						</tr>
					
							<input type="hidden" name="agregardato" value="">
							<input type="hidden" name="IDCliente" value="<%=idcliente%>">
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
		alert("Ud. No tiene autorizaci&oacute;n para este proceso.");
		window.open("dcs_userexpira.asp","_top");
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

