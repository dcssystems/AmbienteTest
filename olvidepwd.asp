<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
''if session("idusuario")<>"" then
	if obtener("agregardato")="1" then
		pwdact=obtener("pwdact")
		pwdnew=obtener("pwdnew")		

		conectar
		pwd=passwordactual()
		if pwd=pwdact then
			actualizapassword pwd
			%>
				<script language="javascript">
					alert("La contrase�a se modific� correctamente.")
					window.close();
				</script>
			<%		
		else
			%>
				<script language="javascript">
					alert("La contrase�a actual ingresada no es v�lida.")
					history.back();
				</script>
			<%
		end if
		desconectar
	end if
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Modificar Contrase�a</title>
			<script language=javascript>
				function agregar()
				{
					if(formula.pwdact.value==""){alert("Debe ingresar su contrase�a actual.");return;}
					if(formula.pwdnew.value==""){alert("Debe ingresar una nueva contrase�a.");return;}
					if(formula.pwdnew.value!=formula.pwdnew2.value){alert("El re-ingreso de la nueva contrase�a no coincide con la nueva contrase�a ingresada.");return;}
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
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
				<form name=formula method=post action="cambiarpwd.asp">
					<table border=0 cellspacing=0 cellpadding=0 width=100%>
					<tr>	
						<td colspan=3>			
							<font size=2 color=#00529B face=Arial><b>&nbsp;<b>Modificar Contrase�a:</b></b></font>
						</td>
					</tr>
					<tr>
						<td><font face=Arial size=2 color=#00529B>&nbsp;&nbsp;Ingrese Contrase�a Actual:</font></td>
						<td colspan=2><input name=pwdact type=password size=18 maxlength=255 value="" style="font-size: xx-small; width: 100px;"></td>
					</tr>
					<tr>	
						<td colspan=3><hr></td>
					</tr>					
					<tr>
						<td><font face=Arial size=2 color=#00529B>&nbsp;&nbsp;Ingrese Nueva Contrase�a:</font></td>
						<td colspan=2><input name=pwdnew type=password size=18 maxlength=255 value="" style="font-size: xx-small; width: 100px;"></td>
					</tr>					
					<tr>
						<td><font face=Arial size=2 color=#00529B>&nbsp;&nbsp;Re-Ingrese Nueva Contrase�a:</font></td>
						<td><input name=pwdnew2 type=password size=18 maxlength=255 value="" style="font-size: xx-small; width: 100px;"></td>
						<td align=right height=40><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>
					</tr>
					</table>
					<input type=hidden name="agregardato" value="">
				</form>	
			</body>
		</html>	
		<%		
''else
%>
<!--<script language="javascript">
	alert("Tiempo Expirado");
	window.close();
</script>-->
<%
''end if
%>
