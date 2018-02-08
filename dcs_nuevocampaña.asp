<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admcrearcampa�a.asp") then
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
				<script language="javascript">
					<%if obtener("agregardato")="1" then%>
					//alert("Se agreg� el usuario correctamente.");
					<%else%>
					//alert("Se modific� el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="dcs_admfacultad.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language="javascript">
					alert("El usuario ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codfacultad<>"" then
					sql="SELECT A.*, " & _
						"B.nombres as Nombreusureg, " & _
						"B.apepaterno as Apepatusureg, " & _
						"B.apematerno as Apematusureg, " & _
						"C.nombres as Nombreusumod, " & _
						"C.apepaterno as Apepatusumod, " & _
						"C.apematerno as Apematusumod " & _
						"FROM campa�a A " & _
						"INNER JOIN usuario B ON B.codusuario=A.usuarioregistra " & _
						"LEFT OUTER JOIN usuario C ON C.codusuario=A.usuariomodifica " & _
						"INNER JOIN Cliente D ON A.IDCliente = D.IDCliente " & _
						"INNER JOIN TipoCampa�a E ON A.IDTipoCampa�a = E.IDTipoCampa�a " & _
						"WHERE A.idcampa�a = " & codfacultad
					consultar sql,RS
					descripcion=rs.Fields("Descripcion")
					codcliente=rs.Fields("IDCliente")		
					fechInicio=rs.Fields("FechaInicio")	
					fechFin=rs.Fields("FechaFin")						
					tipocampana=rs.Fields("IDTipoCampa�a")
					flaghistorico=rs.Fields("FlagHistorico")						
					estado=rs.Fields("Estado")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					RS.Close
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		<head>
			<title><%if codfacultad="" then%>Nuevo <%end if%>Campa�a</title>
			
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->
			
			<script language="javascript" src="scripts/popcalendar.js"></script> 
			<script language="javascript">
				var limpioclave=0;
				<%if codfacultad="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripci�n.");return;}
					if(trim(formula.pagina.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato num�rico.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripci�n.");return;}
					if(trim(formula.pagina.value)==""){alert("Debe asignar un link.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato num�rico.");return;}
					
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
		<body topmargin="0" leftmargin="0" bgcolor="#FFFFFF">
			<table border="0" cellspacing="0" cellpadding="0" width="100%" height="100%">
				<form name="formula" method="post" action="dcs_nuevofacultad.asp">					
					<tr class="fondo-red">	
						<td class="text-withe" colspan="2">			
							<font size="2"><b>&nbsp;<b><%if codfacultad="" then%>Nuevo <%end if%>Campa�a</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr class="fondo-gris" height="25">
						<td class="text-orange label-registra" colspan="2" align="right"><font size="1">Registr�:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modific�:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>
					<tr>
						<td class="text-orange"><font size="2">Cliente:</font></td>
						<td>
							<select name="cliente" style="font-size: xx-small; width: 200px;">
							<%
							sql = "SELECT IDCliente, RazonSocial FROM cliente ORDER BY RazonSocial ASC"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("IDCliente")%>" <% if codcliente<>"" then%><% if RS.fields("IDCliente")=int(codcliente) then%> selected<%end if%><%end if%>><%=RS.Fields("RazonSocial")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange"><font size="2">Tipo de Campa�a:</font></td>
						<td>
							<select name="codgrupofacultad" style="font-size: xx-small; width: 200px;">
							<%
							sql = "SELECT IDTipoCampa�a, Descripcion FROM TipoCampa�a ORDER BY IDTipoCampa�a ASC"
							consultar sql,RS
							Do While Not  RS.EOF
							%>
								<option value="<%=RS.Fields("IDTipoCampa�a")%>" <% if tipocampana<>"" then%><% if RS.fields("IDTipoCampa�a")=int(tipocampana) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
							<%
							RS.MoveNext
							loop
							RS.Close
							%>
							</select>
						</td>
					</tr>
					<!--  class="fondo-gris" -->					
					<tr>
						<td class="text-orange" width="20%"><font  size="2">Descripci�n:</font></td>
						<td><input name="descripcion" type="text" maxlength="200" value="<%=Descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange" width="20%"><font  size="2">Fecha Inicio:</font></td>
						<td><input name="descripcion" type="text" maxlength="200" value="<%=Descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td class="text-orange" width="30%"><font size="2" >Fecha Fin:</font></td>
						<td><input name="pagina" type="text" maxlength="200" value="<%=pagina%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr class="fondo-gris">
						<td class="text-orange" width="30%"><font size="2">Orden:</font></td>
						<td><input name="orden" type="text" maxlength="50" value="<%=orden%>" style="font-size: xx-small; width: 60px; text-align: right"></td>
					</tr>			
					<tr class="fondo-red">					
						<td><font size="2" >&nbsp;</font></td>
						<td align="right" height="40">
							<%if codfacultad="" then%>
							<a href="javascript:agregar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<%else%>
							<a href="javascript:modificar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<%end if%>
							<a href="javascript:window.close();"><i class="logout demo-icon icon-logout">&#xe800;</i></a>&nbsp;
						</td>					
					</tr>

							<input type="hidden" name="agregardato" value="">
							<input type="hidden" name="codfacultad" value="<%=codfacultad%>">
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
		alert("Ud. No tiene autorizaci�n para este proceso.");
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

