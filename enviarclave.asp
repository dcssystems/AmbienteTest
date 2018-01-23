<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if obtener("agregardato")="1" then
	usuario=obtener("usuario")
	email=obtener("email")		

	conectar
		existeusuario=0
		sql="select A.nombres,A.clave,A.activo,A.flagbloqueo,CASE WHEN A.codagencia>0 THEN (select razonsocial from Agencia where codagencia=A.codagencia) WHEN A.codoficina is not null THEN (select codoficina + ' - ' + descripcion from Oficina where codoficina=A.codoficina) ELSE '' END as Entidad from Usuario A where A.usuario='" & usuario & "' and A.correo='" & email & "'"
		consultar sql,RS
		if not RS.EOF then
			existeusuario=1
			nombre=RS.fields("nombres")
			clave=RS.fields("clave")
			activo=RS.fields("activo")
			flagbloqueo=RS.fields("flagbloqueo")
			entidad=RS.fields("entidad")
		else
			existeusuario=0
		end if
		RS.Close
		if existeusuario=0 then
		%>
		<html>
		<body topmargin=0 leftmargin=0 bgcolor="#000000">
				<table border=0 cellspacing=0 cellpadding=0 width=90% height=100% align=center>				
					<tr>	
						<td>			
							<font size=2 color=#FFFFFF face=Arial><b>&nbsp;<b>Olvidé mi contraseña</b></b></font>
						</td>
					</tr>
					<tr>	
						<td><hr></td>
					</tr>						
					<tr>
						<td align=middle><font face=Arial size=2 color=#FF0000>El usuario y el e-mail ingresados no se encuentran en nuestros registros. Comuníquese con la Unidad de Cobranzas BBVA.</font></td>
					</tr>
					<tr>	
						<td><hr></td>
					</tr>					
					</table>			
		</body>
		</html>
		<%
		else
			if activo=1 and flagbloqueo<3 then
			    if entidad<>"" then
			        cadenaentidad=" de la Entidad: <b>" & entidad & "</b> "
			    end if
			    sql="EXEC SP_ENVIARMAIL " & _ 
				    "@profile = 'Cobranzas Notificaciones', " & _ 
				    "@asunto = 'Cobranzas BBVA - Recordatorio de Contraseña', " & _ 
				    "@cuerpo = '<html><body><img border=0 src=" & chr(34) & "https://extranetperu.grupobbva.pe/cobranzacm/imagenes/logo.gif" & chr(34) & "><BR><BR><font face=Arial size=2>Estimado(a) " & nombre & ":<BR><BR>Nos es grato saludarte y hacerte recordar que para acceder a Optimus - Sistema Web de Gestión de Cobranzas BBVA, para el usuario: <b>" & usuario & "</b>" & cadenaentidad & " la contraseña es: <b>" & desencriptar(clave) & "</b>.<BR><BR>Web Site: <a href=https://extranetperu.grupobbva.pe/cobranzacm>https://extranetperu.grupobbva.pe/cobranzacm</a><BR><BR>Saludos,<BR><BR><b>Unidad de Cobranzas BBVA</b></font></body></html>', " & _
				    "@destinatarios = '" & email & "', " & _ 
				    "@copias = '', " & _ 
				    "@copiasocultas = '', " & _ 
				    "@adjuntos = '', " & _ 
				    "@formato='HTML';"
			    conn.execute sql
			end if		
		%>
		<html>
		<body topmargin=0 leftmargin=0 bgcolor="#000000">
				<table border=0 cellspacing=0 cellpadding=0 width=90% height=100% align=center>				
					<tr>	
						<td>			
							<font size=2 color=#FFFFFF face=Arial><b>&nbsp;<b>Olvidé mi contraseña</b></b></font>
						</td>
					</tr>
					<tr>	
						<td><hr></td>
					</tr>						
					<tr>
						<td align=middle><%if activo=1 and flagbloqueo<3 then%><font face=Arial size=2 color=#FFFFFF>En breve recibirás un e-mail con la información de tu cuenta.</font><%else%><%if activo=0 and flagbloqueo<3 then%><font face=Arial size=2 color=#FF0000>El usuario se encuentra inactivo. Comuníquese con la Unidad de Cobranzas BBVA.</font><%else%><font face=Arial size=2 color=#FF0000>El usuario se encuentra bloqueado. Comuníquese con la Unidad de Cobranzas BBVA.</font><%end if%><%end if%></td>
					</tr>
					<tr>	
						<td><hr></td>
					</tr>					
					</table>			
		</body>
		</html>		
		<%
		end if		
	desconectar
else%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Olvidé mi contraseña</title>
			<script language=javascript>
				function enviar()
				{
					if(formula.usuario.value==""){alert("Debe ingresar el usuario.");return;}
					if(formula.email.value==""){alert("Debe ingresar el e-mail.");return;}
					if(!isEmailAddress(formula.email)){alert("Debe ingresar un e-mail válido.");return;}
					document.formula.agregardato.value=1;
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
			<body topmargin=0 leftmargin=0 bgcolor="#000000">
				<form name=formula method=post action="enviarclave.asp">
					<table border=0 cellspacing=0 cellpadding=0 width=90% height=100% align=center>				
					<tr>	
						<td colspan=3>			
							<font size=2 color=#FFFFFF face=Arial><b>&nbsp;<b>Olvidé mi contraseña</b></b></font>
						</td>
					</tr>
					<tr>	
						<td colspan=3><hr></td>
					</tr>						
					<tr>
						<td width=20%><font face=Arial size=2 color=#FFFFFF>&nbsp;&nbsp;Usuario:</font></td>
						<td><input type="text" name="usuario" value="" style="font-size:xx-small; width:100px; height:18px;"></td>
						<td>&nbsp;</td>
					</tr>					
					<tr>
						<td width="20%"><font face=Arial size=2 color=#FFFFFF>&nbsp;&nbsp;e-mail:</font></td>
						<td><input type="text" name="email" value="" style="font-size:xx-small; width:180px; height:18px;"></td>
						<td><a href="javascript:enviar();" style="text-decoration:none;"><img src="imagenes/btenviar.png" alt="Enviar" title="Enviar" border=0></a></td>
					</tr>
					<tr>	
						<td colspan=3><hr></td>
					</tr>					
					</table>
					<input type="hidden" name="agregardato" value="">
				</form>	
			</body>
		</html>
<%
end if	
%>
