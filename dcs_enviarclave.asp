<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if obtener("agregardato")="1" then
	usuario=obtener("usuario")
	email=obtener("email")		

	conectar
		existeusuario=0
		sql="SELECT usr.nombres, usr.clave, usr.activo, usr.flagbloqueo FROM Usuario usr WHERE usr.usuario='" & usuario & "' AND usr.correo='" & email & "'"
		consultar sql,RS
		if not RS.EOF then
			existeusuario=1
			nombre=RS.fields("nombres")
			clave=RS.fields("clave")
			activo=RS.fields("activo")
			flagbloqueo=RS.fields("flagbloqueo")
			'entidad=RS.fields("entidad")
		else
			existeusuario=0
		end if
		RS.Close
		if existeusuario=0 then
		%>
		<div class="row">
			<div class="col-sm-12">
				<h5 class="text-center">Olvidé mi contraseña</h5>				
				<div class="alert alert-danger alert-dismissable fade in">					
					<a href="#" class="close" data-dismiss="alert" aria-label="close">&times;</a>
					<strong>Cuidado!</strong> El usuario y el e-mail ingresados no se encuentran en nuestros registros. Comuníquese con el Administrador.
				</div>
			</div>
		</div>
		<%
		else
			if activo=1 and flagbloqueo<3 then
			    if entidad<>"" then
			        cadenaentidad=" de la Entidad: <b>" & entidad & "</b> "
			    end if
			    sql="EXEC SP_ENVIARMAIL " & _ 
				    "@profile = 'CRM - DIRCON | Notificaciones', " & _ 
				    "@asunto = 'CRM DIRCON - Recordatorio de Contraseña', " & _ 
				    "@cuerpo = '<html><body><img border=0 src=" & chr(34) & "https://extranetperu.grupobbva.pe/cobranzacm/imagenes/logo.gif" & chr(34) & "><BR><BR><font face=Arial size=2>Estimado(a) " & nombre & ":<BR><BR>Nos es grato saludarte y hacerte recordar que para acceder a Optimus - Sistema CRM DIRCON, para el usuario: <b>" & usuario & "</b>" & cadenaentidad & " la contraseña es: <b>" & desencriptar(clave) & "</b>.<BR><BR>Web Site: <a href=https://extranetperu.grupobbva.pe/cobranzacm>https://extranetperu.grupobbva.pe/cobranzacm</a><BR><BR>Saludos,<BR><BR><b>Equipo de Desarrollo | DIRCON</b></font></body></html>', " & _
				    "@destinatarios = '" & email & "', " & _ 
				    "@copias = '', " & _ 
				    "@copiasocultas = '', " & _ 
				    "@adjuntos = '', " & _ 
				    "@formato='HTML';"
			    conn.execute sql
			end if		
		%>
		<div class="row">
			<div class="col-sm-12">
				<h3 class="text-center">Olvidé mi contraseña</h3>				
				<%if activo=1 and flagbloqueo<3 then%>
					<p class="alert alert-success">En breve recibirás un e-mail con la información de tu cuenta.</p>
				<%else%>
					<%if activo=0 and flagbloqueo<3 then%>
						<p class="alert alert-danger">El usuario se encuentra inactivo. Comuníquese con el administrador.</p>
					<%else%>
						<p class="alert alert-danger">El usuario se encuentra bloqueado. Comuníquese con el administrador.</p>
					<%end if%>
				<%end if%>				
			</div>
		</div>
		<%end if
		desconectar
else%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			:)
		</html>
<%
end if	
%>
