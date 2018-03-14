<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->
<!DOCTYPE html>
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
			<input type="button" value="llamar" onclick="llamar('51','944267001');return false;">
			<input type="button" value="Obtener" onclick="getValuesIframe();" />
			
			<!--MDDiaLCodE-->
			<!--MDPhonENumbeR-->
			<input name="envio" id="envio" type="button" value="TEST" onclick="pasarmensaje();">
			
			<input type="text" name="contestoZoiper" id="contestoZoiper" value="0" onchange="alert('ok contestaste');return false;"/>
		
		</form>
		
		
		<script language="javascript">
		
			function pasarmensaje(){
				//var frame = document.getElementById('interloc'); 
				//window.parent.frames[1].contentWindow.postMessage("jajajaja", "http://192.168.1.7");
				var frame = window.parent.frames[1]; 
				frame.contentWindow.postMessage("jajajaja", "http://192.168.1.5"); 
			}
			
		
		function displayMessage (evt) {
			var message;
			alert(evt.origin);
			if (evt.origin !== "http://192.168.1.5") {
				message = "You are not worthy";
			}
			else {
				message = "I got " + evt.data + " from " + evt.origin;
			}
			//document.getElementById("received-message").innerHTML = message;
			alert(message);
		}

		if (window.addEventListener) {
			// For standards-compliant web browsers
			window.addEventListener("message", displayMessage, false);
		}
		else {
			window.attachEvent("onmessage", displayMessage);
		}
		
		
		
		
			/*function llamar(codeCountry,phoneNumber)
			{
				var xhttp = new XMLHttpRequest();
				xhttp.onreadystatechange = function() {
					if (this.readyState == 4 && this.status == 200) {
						document.getElementById("demo").innerHTML = this.responseText;
					}
				};
				xhttp.open("POST", "http://192.168.1.5/interlocutor.php", true);
				xhttp.send("codeCountry=" + codeCountry + "&phoneNumber=" + phoneNumber); 
			}*/
		</script>
		
		<!--<iframe src="http://192.168.1.5/interlocutor.php" height="90%" width="100%" id="interloc" name="interloc"></iframe>
		
		<hr>
		iframe src="http://192.168.1.5/agent"  height="80%" width="100%" id="autodial" name="autodial"></iframe-->
			

	</body>
</html>