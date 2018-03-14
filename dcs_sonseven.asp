<html>
	<body>			
		
		<input type="submit" name="SUBMIT" value="Login"  />
		<input type="button" value="llamar" onclick="llamar('51','944267001');return false;">
		<input type="button" value="Obtener" onclick="getValuesIframe();" />
		
		
		<input name="envio" id="envio" type="button" value="TEST" onclick="pasarmensaje();">
		
		<input type="text" name="contestoZoiper" id="contestoZoiper" value="0" onchange="alert('ok contestaste');return false;"/>
		
		
		
		
		<center>
			Aqui estamos
		</center>
		
	</body>
		<script language="javascript">
		function displayMessage (evt) {
				var message;
				alert(evt.origin);
				if (evt.origin !== "http://192.168.1.5") {
					message = "You are not worthy - (hijo del siete)";
				}
				else {
					message = "I got " + evt.data + " from " + evt.origin;
				}
				//document.getElementById("received-message").innerHTML = message;
				alert(message);
				top.mensajefinal("jajaja Regrese a 7");
			}

			if (window.addEventListener) {
				// For standards-compliant web browsers
				window.addEventListener("message", displayMessage, false);
			}
			else {
				window.attachEvent("onmessage", displayMessage);
			}
		
	
	</script>
</html>










