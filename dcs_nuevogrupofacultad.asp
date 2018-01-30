<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admgrupofacultad.asp") then
	buscador=obtener("buscador")	
	codgrupofacultad=obtener("codgrupofacultad")
		if obtener("agregardato")<>"" then
		descripcion=obtener("descripcion")
		orden=obtener("orden")
		if not isNumeric(orden) then
			orden="0"
		end if	
		orden=int(orden)
										
									
			existegrupofacultad=0
			
			if codgrupofacultad<>"" then
			sql="select count(*) from grupofacultad where descripcion='" & descripcion & "' and codgrupofacultad<>" & codgrupofacultad 
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
					sql="update grupofacultad set descripcion='" & descripcion & "',orden=" & orden & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codgrupofacultad=" & codgrupofacultad
				end if
				''Response.Write sql
				conn.execute sql
									
				%>
				<script language="javascript">
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="dcs_admgrupofacultad.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
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
			if codgrupofacultad<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from grupofacultad A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where a.codgrupofacultad = " & codgrupofacultad
					consultar sql,RS
					descripcion=rs.Fields("descripcion")	
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
			<head>
			<title><%if codgrupofacultad="" then%>Nuevo <%end if%>Grupo Facultad</title>
			
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->
			
			<script language='javascript' src="scripts/popcalendar_cobcm.js"></script> 
			<script language="javascript">
				var limpioclave=0;
				<%if codgrupofacultad="" then%>
				function agregar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
					if(isNaN(trim(formula.orden.value.replace(",","")))){alert("El orden debe ser un dato numérico.");return;}
																		
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.descripcion.value)==""){alert("Debe ingresar una Descripción.");return;}
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
					<form name="formula" method="post" action="dcs_nuevogrupofacultad.asp">					
						<tr class="fondo-red">	
							<td class="text-withe" colspan="2">			
								<font size="2" ><b>&nbsp;<b><%if codgrupofacultad="" then%>Nuevo <%end if%>Grupo Privilegio</b></b></font>
							</td>
						</tr>
					<%if fechaReg<>"" then%>
					<tr height="20">
						<td class="text-orange label-registra" colspan="2" align="right"><font size="1">Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<tr>
						<td class="text-orange" width="20%"><font size="2">&nbsp;&nbsp;Descripción:</font></td>
						<td><input name="descripcion" type="text" maxlength="200" value="<%=descripcion%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td class="text-orange" width="20%"><font size="2">&nbsp;&nbsp;Orden:</font></td>
						<td><input name="orden" type="text" maxlength="50" value="<%=orden%>" style="font-size: xx-small; width: 60px; text-align: right"></td>
					</tr>			
					<tr class="fondo-red">					
						<td><font size="2" >&nbsp;</font></td>
						<td align="right" height="40">
							<%if codgrupofacultad="" then%>
							<a href="javascript:agregar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<%else%>
							<a href="javascript:modificar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;
							<%end if%>
							<a href="javascript:window.close();"><i class="logout demo-icon icon-logout">&#xe800;</i></a>&nbsp;
						</td>					
					</tr>
					
							<input type="hidden" name="agregardato" value="">
							<input type="hidden" name="codgrupofacultad" value="<%=codgrupofacultad%>">
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

