<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp--> 
<%
	
	
%>
<!doctype html>
<html lang="en">
<head>
	<meta charset="UTF-8">
	<title>Abuelo 7</title>
</head>
<body>

	<body topmargin="0" leftmargin="0">
	<form name="formula" >
		<input name="login" id="login" type="button" value="Login" onclick="enviardatosp5('LOGIN');">
		
		<input name="llamar" id="llamar" type="button" value="Llamar" onclick="enviardatosp5('LLAMAR');" disabled>
	</form>				
		<center>
			<iframe src="http://192.168.1.5/p5.php" id="mandophp" name="mandophp" height="90%" width="90%"  style="height: 800px;" ></iframe>
		</center>
		
		<script language="javascript">		
			async function enviardatosp5(cadena)			
			{
				//var frame = document.getElementById('interloc'); 
				//window.parent.frames[1].contentWindow.postMessage("jajajaja", "http://192.168.1.7");
				var frameV = window.parent.frames[0].frames[1]; 
				//postMessage(mensaje, ip a donde mando el mensaje);				
				//
				switch(cadena){
					case "LOGIN":
						frameV.postMessage("E-LOGIN", "http://192.168.1.5");
						await sleep(250);
						frameV.postMessage("D-USUARIO|<%=DUSUARIO%>", "http://192.168.1.5");
						//await sleep(1);
						frameV.postMessage("D-PASSWORD|<%=DPASSWORD%>", "http://192.168.1.5");
						//await sleep(1);
						frameV.postMessage("D-TELFLOGIN|<%=DTELFLOGIN%>", "http://192.168.1.5");
						//await sleep(1);
						frameV.postMessage("D-TELFPASSWORD|<%=DTELFPASSWORD%>", "http://192.168.1.5");
						await sleep(250);
						frameV.postMessage("D-CAMP|18021801", "http://192.168.1.5");	
						await sleep(250);
						frameV.postMessage("F-LOGIN", "http://192.168.1.5");							
					break;
					
					case "LLAMAR":
						frameV.postMessage("E-LLAMAR", "http://192.168.1.5");
						//await sleep(1);
						frameV.postMessage("D-CODPAIS|<%=DCODPAIS%>", "http://192.168.1.5");
						//await sleep(1);
						frameV.postMessage("D-TELEFONO|981603575", "http://192.168.1.5");						
						await sleep(250);
						frameV.postMessage("F-LLAMAR", "http://192.168.1.5");							
					break;				
					
				}				 
			}			
			
			function sleep(ms)
			{
			  return new Promise(resolve => setTimeout(resolve, ms));
			}
			
			function respuesta_peticion(cadena)			
			{
				switch(cadena)
				{
					case "D-LOGINOK":
						document.formula.login.disabled = true;
						document.formula.login.value = "LOGIN OK";
						document.formula.llamar.disabled = false;
				}			
				 
			}
		</script>
</body>
</html>