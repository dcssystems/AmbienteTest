<html>

	<body>
			
		
		<input type="hidden" name="my-message" id="my-message">
		
		<iframe src="http://192.168.1.7/ambientetest/dcs_logingoautodial.asp"  height="20%" width="100%"   id="mandocrm" name="mandocrm"></iframe>
			<hr>
		<iframe src="http://192.168.1.5/interlocutor.php" height="80%" width="100%" id="interloc" name="interloc"></iframe>
			<hr>
		<!--<iframe src="http://192.168.1.5/agent"  height="50%" width="100%" id="autodial" name="autodial"></iframe>
			<hr>
			
			<iframe height="10%" width="100%" id="mandocrm1" name="mandocrm1"></iframe-->
		
	</body>
		<script language="javascript">
		
		function hola(){
		alert("hola");
		}
		
		/*window.onload = function () {
			var iframeWin = document.getElementById("interloc").contentWindow;
			var btnEnvio  = document.getElementById("envio");
			var myMessage = document.getElementById("my-message").value = "Hello World!!";
			
			//myMessage.select();
			
			btnEnvio.onsubmit = function () {
				iframeWin.postMessage(myMessage.value, "http://192.168.1.5");
				return false;
			};

		};*/
	
	</script>
</html>