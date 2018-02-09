<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admCampañacampo.asp") then
		buscador=obtener("buscador")	
		idcliente=obtener("IDCliente")
		if obtener("agregardato")<>"" then
			RazonSocial=obtener("RazonSocial")
			RUC=obtener("RUC")
			Direccion=obtener("Direccion")
			Telefono=obtener("Telefono")
			Email=obtener("Email")
			if obtener("Activo")<>"" then Activo=1 else Activo=0 end if											
									
			existeCliente=0
			
			if idcliente<>"" then
			sql="select count(*) from Cliente where RazonSocial = '" & RazonSocial & "' and IDCliente <>" & idcliente 
			else
			sql="select count(*) from Cliente where RazonSocial = '" & RazonSocial & "'"
			end if
			consultar sql,RS
			existeCliente=RS.Fields(0)
			RS.Close			
			if existeCliente=0 then			


				if obtener("agregardato")="1" then		
					sql="insert into Cliente (RazonSocial,RUC,Direccion,Telefono,Email,Activo,UsuarioRegistra,FechaRegistra)  values ('" & RazonSocial & "','" & RUC & "','" & Direccion & "','" & Telefono & "','" & Email & "', " & Activo & "," & session("codusuario") & ",getdate())"
				else
					sql="update Cliente set RazonSocial='" & RazonSocial & "',RUC='" & RUC & "',Direccion='" & Direccion & "',Telefono='" & Telefono & "',Email='" & Email & "',Activo=" & Activo & ",UsuarioModifica=" & session("codusuario") & ",fechamodifica=getdate() where IDCliente = " & idcliente
				end if
				'Response.Write sql
				conn.execute sql
									
				%>
				<script language="javascript">
					<%if obtener("agregardato")="1" then%>
					//alert("Se agreg&oacute; el Cliente correctamente.");
					<%else%>
					//alert("Se modific&oacute; el Cliente correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="dcs_admCliente.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language="javascript">
					alert("El Cliente ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if idcliente<>"" then
					sql="SELECT a.IDCampañaCampo, " &_
					"a.GlosaCampo, " &_
					"a.Nivel, " &_
					"a.TipoCampo, " &_
					"a.FlagNroDocumento, " &_
					"a.Visible, " &_
					"d.Descripcion AS 'Descampaña', " &_
					"A.fecharegistra, " &_
					"B.nombres AS nombresReg, " &_
					"B.apepaterno AS apepaternoReg, " &_
					"B.apematerno AS apematernoReg, " &_
					"A.fechamodifica, " &_
					"C.nombres AS nombresMod, " &_
					"C.apepaterno AS apepaternoMod, " &_
					"C.apematerno AS apematernoMod " &_
					"FROM Campaña_Campo a  " &_
					"LEFT OUTER JOIN Usuario B ON A.usuarioregistra=B.codusuario " &_
					"LEFT OUTER JOIN Usuario C ON A.usuariomodifica=C.codusuario " &_
					"INNER JOIN TipoCampaña d ON a.IDTipoCampaña = d.IDTipoCampaña " &_
					"WHERE IDCampañaCampo = " & idcliente
					consultar sql,RS
					idcampanacampo   = rs.Fields("IDCampañaCampo")	
					glosacampo       = rs.Fields("GlosaCampo")
					nivel            = rs.Fields("Nivel")	
					tipoCampo        = rs.Fields("TipoCampo")
					FlagNroDocumento = rs.Fields("FlagNroDocumento")
					visible          = rs.Fields("Visible")	
					descampana       = rs.Fields("Descampaña")					
					FechaRegistra    = RS.Fields("FechaRegistra")
					usuarioReg       = iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					FechaModifica    = RS.Fields("FechaModifica")
					usuarioMod       = iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
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
					if(trim(formula.RazonSocial.value)==""){alert("Debe ingresar una Raz&oacute;n Social.");return;}
					if(trim(formula.RUC.value)==""){alert("Debe ingresar un RUC.");return;}
					if(!isEmailAddress(formula.Email)){alert("Debe ingresar un e-Mail v&aacute;lido.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
				    if(trim(formula.RazonSocial.value)==""){alert("Debe ingresar una Raz&oacute;n Social.");return;}
					if(trim(formula.RUC.value)==""){alert("Debe ingresar un RUC.");return;}
					if(!isEmailAddress(formula.Email)){alert("Debe ingresar un e-Mail v&aacute;lido.");return;}
					
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
						<%if fechaRegistra<>"" then%>
						<tr height="20">
							<td class="text-orange label-registra" colspan="2" align="right"><font size="1">Registr&oacute;:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=FechaRegistra%></b>
							<%if FechaModifica<>"" then%><BR>Modific&oacute;:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
							</font></td>
						</tr>
						<%end if%>	
						<!-- SECCION DEL FORMULARIO PARA CLIENTE -->						
						<tr class="fondo-gris">
							<td class="text-orange" width="30%"> Raz&oacute;n Social:</td>
							<td><input name="RazonSocial" type="text" maxlength="200" value="<%=RazonSocial%>" style="font-size: xx-small; width: 200px;"></td>
						</tr>
						<tr>
							<td class="text-orange" width="30%"><font size="2"> Documento de Identidad: </font></td>
							<td><input name="RUC" type="text" maxlength="200" value="<%=RUC%>" style="font-size: xx-small; width: 200px; text-align: left"></td>
						</tr>
						<tr class="fondo-gris">
							<td class="text-orange" width="30%"><font size="2"> Direcci&oacute;n:</font></td>
							<td><input name="Direccion" type="text" maxlength="200" value="<%=Direccion%>" style="font-size: xx-small; width: 200px;"></td>
						</tr>
						<tr>
							<td class="text-orange" width="30%"><font size="2"> Tel&eacute;fono:</font></td>
							<td><input name="Telefono" type="text" maxlength="200" value="<%=Telefono%>" style="font-size: xx-small; width: 200px; text-align: left"></td>
						</tr>
						<tr class="fondo-gris">
							<td class="text-orange" width="30%"><font size="2"> Email:</font></td>
							<td><input name="Email" type="text" maxlength="200" value="<%=Email%>" style="font-size: xx-small; width: 200px;"></td>
						</tr>
						<tr>
							<td class="text-orange" width="30%"><font size="2"> Activo:</font></td>
							<td><input type=checkbox name="activo"  <%if activo=1 then%> checked<%end if%>></td>
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

