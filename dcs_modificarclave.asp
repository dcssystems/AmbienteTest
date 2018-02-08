<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then

	if obtener("agregardato")="1" then
		pwdact=obtener("pwdact")
		pwdnew=obtener("pwdnew")		

		conectar
		pwd=passwordactual()
		if pwd=pwdact then
			actualizapassword pwd
			%>
				<script language="javascript">
					alert("La contraseña se modificó correctamente.")
					window.close();
				</script>
			<%		
		else
			%>
				<script language="javascript">
					alert("La contraseña actual ingresada no es válida.")
					history.back();
				</script>
			<%
		end if
		desconectar
	end if
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		<head>
			<title>Modificar Contraseña</title>
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->
			<script language="javascript">
				function agregar()
				{
					if(formula.pwdact.value==""){alert("Debe ingresar su contraseña actual.");return;}
					if(formula.pwdnew.value==""){alert("Debe ingresar una nueva contraseña.");return;}
					if(formula.pwdnew.value!=formula.pwdnew2.value){alert("El re-ingreso de la nueva contraseña no coincide con la nueva contraseña ingresada.");return;}
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
			</script>
		</head>
			<body topmargin="0" leftmargin="0" style="margin-top: 10px;">
				<form name="formula" method="post" action="dcs_modificarclave.asp">
					<table border="0" cellspacing="0" cellpadding="0" width="100%">
					<tr>	
						<td class="text-orange" colspan="3">			
							<font size="2"><b>&nbsp;<b>Modificar Contraseña:</b></b></font>
						</td>
					</tr>
					<tr>
						<td class="text-orange">
							<font size="2">&nbsp;&nbsp;Ingrese Contraseña Actual:</font>
						</td>
						<td class="text-orange" colspan="2">
							<input name="pwdact" type="password" size="18" maxlength="255" value="" style="font-size: xx-small; width: 100px;"></td>
					</tr>
					<tr>	
						<td colspan=3><hr></td>
					</tr>					
					<tr>
						<td class="text-orange">
							<font size=2>&nbsp;&nbsp;Ingrese Nueva Contraseña:</font>
						</td>
						<td colspan=2>
							<input name="pwdnew" type="password" size="18" maxlength="255" value="" style="font-size: xx-small; width: 100px;">
						</td>
					</tr>					
					<tr>
						<td class="text-orange">
							<font size=2>&nbsp;&nbsp;Re-Ingrese Nueva Contraseña:</font>
						</td>
						<td class="text-orange"><input name="pwdnew2" type="password" size="18" maxlength="255" value="" style="font-size: xx-small; width: 100px;"></td>
						<td align="right" height="40"><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<a href="javascript:window.close();"><img src="imagenes/salida.gif" border="0" alt="Salir" title="Salir"></a>&nbsp;</td>
					</tr>
					</table>
					<input type="hidden" name="agregardato" value="">
				</form>	
			</body>
		</html>	
		<%		
else
%>
<script language="javascript">
    alert("Tiempo Expirado");
    //window.open("index.html","_top");
    window.open("index.html", "sistema");
    window.close();
</script>
<%
end if
%>
