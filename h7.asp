<!doctype html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Hijo 7</title>
</head>
<body>

		<script language="javascript">
		
			function displayMessage (evt) 
			{
				
				if (evt.origin !== "http://192.168.1.5") 
				{
					var message;				
					message = "NO SE CUENTA CON AUTORIZACIÓN PARA REALIZAR ESTA OPERACIÓN.";
					alert(message);
				}
				else 
				{
					top.respuesta_peticion(evt.data);	
				}
				//document.getElementById("received-message").innerHTML = message;
				
			}

			if (window.addEventListener) 
			{
				// For standards-compliant web browsers
				window.addEventListener("message", displayMessage, false);
			}
			else 
			{
				window.attachEvent("onmessage", displayMessage);
			}
			
		</script>	
</body>
</html>