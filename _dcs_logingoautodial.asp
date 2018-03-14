<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->
<!DOCTYPE html>
<script language="javascript">
window.name="abuelo";
</script>
<html lang="es">
	<head>
		<title>Testing Calling</title>
		<meta charset="ISO-8859-1"> 
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet"  href="assets/bootstrap/dist/css/bootstrap.css"/>
	</head>
		
	<body topmargin="0" leftmargin="0"><!--onload="inicio();"-->
	
		<form name="vicidial_form" id="vicidial_form" action="http://192.168.1.5/interlocutor.php" method="post" target="interloc">
			<input type="hidden" name="DB" id="DB" value="" />
			<input type="hidden" name="JS_browser_height" id="JS_browser_height" value="" />
			<input type="hidden" name="JS_browser_width" id="JS_browser_width" value="" />
			
			<input name="relogin" id="relogin" type="hidden" value="NO" />
			<input name="VD_login" id="VD_login" type="hidden" value="Sistemas2" />
			<input name="phone_login" type="hidden" value="598006" />
			<input name="phone_pass"  type="hidden" value="1Aaq9e3wLJ" />
			<input name="VD_pass" id="VD_pass" type="hidden" value="MkYGSDeI5" />
			<input name="VD_campaign" id="VD_campaign" type="hidden" value="18021801" />
			
			
			<input type="submit" name="SUBMIT" value="Login"  />
			<input type="button" value="llamar" onclick="window['interloc'].llamar('51','981603575');">
			<input type="button" value="Obtener" onclick="getValuesIframe();" />
			
			<!--MDDiaLCodE-->
			<!--MDPhonENumbeR-->
		
			<input type="text" name="contestoZoiper" id="contestoZoiper" value="0" onchange="alert('ok contestaste');return false;"/>
		
		</form>
		
		<iframe id="interloc" name="interloc" height="100%" width="100%" ></iframe><!--style="visibility:hidden;"-->

	</body>
</html>