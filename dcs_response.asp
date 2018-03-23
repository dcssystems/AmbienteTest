<!doctype html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Respopnse ASP</title>
</head>
<body>
		dcs_response.asp
		
	
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
				top.mensajefinal(evt.data);
			}

			if (window.addEventListener) {
				// For standards-compliant web browsers
				window.addEventListener("message", displayMessage, false);
			}
			else {
				window.attachEvent("onmessage", displayMessage);
			}
	</script>
	
</body>
</html>


