<!doctype html>
<html lang="es">
<head>
	<meta charset="UTF-8">
	<title>Control Calling ASP</title>
</head>
<body>
	<form name="form-calling" id="form-calling" action="http://192.168.1.5/mediator.php" method="post" target="mediator">
		<label>Estados:</label><br/>
		<label>Paso #1:</label><input name="stepone" id="stepone" value="0" /><br/>
		<label>Paso #2:</label><input name="steptwo" id="steptwo" value="0" /><br/>
		<label>Paso #3:</label><input name="stepthree" id="stepthree" value="0" /><br/>
		<label>Paso #4:</label><input name="stepfour" id="stepfour" value="0" /><br/>
		<label>Paso #5:</label><input name="stepfive" id="stepfive" value="0" /><br/>
		<label>Paso #6:</label><input name="stepsix" id="stepsix" value="0" /><br/>
		<label>Paso #7:</label><input name="stepseven" id="stepseven" value="0" /><br/>		
		
		<input type="hidden" name="DB" id="DB" value="" />
		<input type="hidden" name="JS_browser_height" id="JS_browser_height" value="" />
		<input type="hidden" name="JS_browser_width" id="JS_browser_width" value="" />
		
		<input name="relogin" id="relogin" type="hidden" value="NO" />
		<input name="VD_login" id="VD_login" type="hidden" value="Sistemas2" />
		<input name="phone_login" type="hidden" value="598006" />
		<input name="phone_pass"  type="hidden" value="1Aaq9e3wLJ" />
		<input name="VD_pass" id="VD_pass" type="hidden" value="MkYGSDeI5" />
		<input name="VD_campaign" id="VD_campaign" type="hidden" value="18021801" />
		
		<input name="MDDiaLCodE" id="MDDiaLCodE" type="hidden" value="51" />
		<input name="MDPhonENumbeR1" id="MDPhonENumbeR1" type="hidden" value="981603575" />
		<input name="MDPhonENumbeR2" id="MDPhonENumbeR2" type="hidden" value="944267001" />
		<input name="MDPhonENumbeR3" id="MDPhonENumbeR3" type="hidden" value="963153879" />
		
		
		<input type="submit" name="SUBMIT" value="Login"  />
		<!--input type="button" value="llamar" onclick="llamar('51','944267001');return false;">
		<input type="button" value="Obtener" onclick="getValuesIframe();" /-->
		
		<!--MDDiaLCodE-->
		<!--MDPhonENumbeR-->
		
		<!-- input name="envio" id="envio" type="button" value="TEST" onclick="pasarmensaje();">
		<input type="text" name="contestoZoiper" id="contestoZoiper" value="0" onchange="alert('ok contestaste');return false;"/ -->		
	</form>
	
	<iframe src="http://192.168.1.5/mediator.php" height="90%" width="100%" id="mediator" name="mediator" style="height: 800px;"></iframe>
		




	<script>
	
		function mensajefinal(smf){
			//alert(smf);
			document.getElementById("stepone").value = smf;
			
		}
		
	</script>
	
</body>
</html>