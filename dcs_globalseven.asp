<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->
<!DOCTYPE html>
<html lang="es">
	<head>
		<title>global 7</title>
		<meta charset="ISO-8859-1"> 
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet"  href="assets/bootstrap/dist/css/bootstrap.css"/>
	</head>
		
	<body topmargin="0" leftmargin="0">
		<input name="btnlogin" id="btnlogin" type="button" value="Login" onclick="pasarmensaje();">
		
		<form name="conect" id="conect" method="POST" >
			<label>Pase 1: </label> <input type="text" name="paseuno" id="paseuno" value="" /><br/>
			<label>Pase 2: </label> <input type="text" name="pasedos" id="pasedos" value="" /><br/>
			<label>Pase 3: </label> <input type="text" name="pasetres" id="pasetres" value="" /><br/>
			<label>Pase 4: </label> <input type="text" name="pasecuatro" id="pasecuatro" value="" /><br/>
			<label>Pase 5: </label> <input type="text" name="pasecinco" id="pasecinco" value="" /><br/>
		 				
			<input type="hidden" name="DB" id="DB" value="" />
			<input type="hidden" name="JS_browser_height" id="JS_browser_height" value="" />
			<input type="hidden" name="JS_browser_width" id="JS_browser_width" value="" />
			
			<input name="relogin" id="relogin" type="hidden" value="NO" />
			<input name="VD_login" id="VD_login" type="hidden" value="Sistemas2" />
			<input name="phone_login" id="phone_login" type="hidden" value="598006" />
			<input name="phone_pass" id="phone_pass"  type="hidden" value="1Aaq9e3wLJ" /><!-- 1Aaq9e3wLJ -->
			<input name="VD_pass" id="VD_pass" type="hidden" value="MkYGSDeI5" />
			<input name="VD_campaign" id="VD_campaign" type="hidden" value="18021801" />
		
		</form>
		
		<center>
			<iframe src="http://192.168.1.5/dcs_globalfive.php" id="mandophp" name="mandophp" height="90%" width="90%"  style="height: 800px;" ></iframe>
		</center>
		
		<script language="javascript">
			function getHiddenValues()
			{
				const DB 				= document.getElementById('DB').value;
				const JS_browser_height = document.getElementById('JS_browser_height').value;
				const JS_browser_width 	= document.getElementById('JS_browser_width').value;
				const relogin			= document.getElementById('relogin').value;
				const VD_login 			= document.getElementById('VD_login').value;
				const phone_login 		= document.getElementById('phone_login').value;
				const phone_pass 		= document.getElementById('phone_pass').value;
				const VD_pass 			= document.getElementById('VD_pass').value;
				const VD_campaign	 	= document.getElementById('VD_campaign').value;
				
				var cadenaUno = "?DB="+DB+"&JS_browser_height="+JS_browser_height+"&JS_browser_width="+JS_browser_width+"&relogin="+relogin+"&VD_login="+VD_login;
				var cadenaDos = "&phone_login="+phone_login+"&phone_pass="+phone_pass+"&VD_pass="+VD_pass+"&VD_campaign="+VD_campaign
				var cadena    = cadenaUno.concat(cadenaDos);
				document.getElementById("paseuno").value = 1;
				return cadena;				
			}
		
			
			function pasarmensaje(){
				//var frame = document.getElementById('interloc'); 
				//window.parent.frames[1].contentWindow.postMessage("jajajaja", "http://192.168.1.7");
				var frame = window.parent.frames[0]; 
				//postMessage(mensaje, ip a donde mando el mensaje);				
				//get element hidden
				var hiddenValues = getHiddenValues();
				frame.postMessage(hiddenValues, "http://192.168.1.5"); 
			}		
		
			function displayMessage (evt) {
				var message;
				alert(evt.origin);
				if (evt.origin !== "http://192.168.1.5") {
					message = "You are not worthy- (Papa del siete) ";
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
						
			function mensajefinal(smf){
				alert(smf);
			}
		</script>
		

	</body>
</html>