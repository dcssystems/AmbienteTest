<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->
<!DOCTYPE html>
<html lang="es">
	<head>
		<title>Testing Calling</title>
		<meta name="viewport" content="width=device-width, initial-scale=1">
		<link rel="stylesheet"  href="assets/bootstrap/dist/css/bootstrap.css"/>
	</head>
		
	<body topmargin="0" leftmargin="0"><!--onload="inicio();"-->
	
		<form name="vicidial_form" id="vicidial_form" action="http://192.168.1.5/login/interlocutor.php" method="post" target="t_new">
			<input type="hidden" name="DB" id="DB" value="" />
			<input type="hidden" name="JS_browser_height" id="JS_browser_height" value="" />
			<input type="hidden" name="JS_browser_width" id="JS_browser_width" value="" />
			
			<input name="relogin" id="relogin" type="hidden" value="NO" />
			<input name="VD_login" id="VD_login" type="hidden" value="Sistemas2" />
			<input name="phone_login" type="hidden" value="598006" />
			<input name="phone_pass"  type="hidden" value="1Aaq9e3wLJ" />
			<input name="VD_pass" id="VD_pass" type="hidden" value="MkYGSDeI5" />
			<input name="VD_campaign" id="VD_campaign" type="hidden" value="10021801" />
			
			
			<input type="submit" name="SUBMIT" value="Login"  />
			<input type="button" value="llamar" onclick="llamar();">
			<input type="button" value="Obtener" onclick="getValuesIframe();" />
			
			<!--MDDiaLCodE-->
			<!--MDPhonENumbeR-->
		
		</form>
		
		<iframe id="interloc" name="interloc" style="height:100%; width:100%;" ></iframe><!--style="visibility:hidden;"-->

		
		<!-- jQuery first, then Tether, then Bootstrap JS. -->
		<script src="https://code.jquery.com/jquery-3.1.1.slim.min.js" integrity="sha384-A7FZj7v+d/sdmMqp/nOQwliLvUsJfDHW+k9Omg/a/EheAdgtzNs3hpfag6Ed950n" crossorigin="anonymous"></script>
		<script src="https://cdnjs.cloudflare.com/ajax/libs/tether/1.4.0/js/tether.min.js" integrity="sha384-DztdAPBWPRXSA/3eYEEUWrWCy7G5KFbe8fFjk5JAIxUYHKkDx6Qin1DkWx51bBrb" crossorigin="anonymous"></script>
		<script src="https://maxcdn.bootstrapcdn.com/bootstrap/4.0.0-alpha.6/js/bootstrap.min.js" integrity="sha384-vBWWzlZJ8ea9aCX4pEW3rVHjgjt7zpkNpZk+02D9phzyeVkE+jo0ieGizqPLForn" crossorigin="anonymous"></script>
  
		
		<script language="javascript">
		function llamar(valueContent)
		{
			var phone = $(valueContent).html();
			alert(phone);
			
			
			
			//window["t_new"].getElementById("NewManualDialTableID").MDDiaLCodE.value="51";
		}
		
		function codeCountry(){
			var myIFrameOne = document.querySelector('#t_new');
			//console.log(myIFrameOnownerDocumente.ownerDocument);
			console.log("1) ownerDocument " + myIFrameOne.ownerDocument.name());
			//console.log("2) contentDocument " + myIFrameOne.contentDocument.getElementById("NewManualDialTableID"));
			//console.log("3) contentWindows " + myIFrameOne.contentWindow.getElementById("NewManualDialTableID"));
			console.log(document.getElementById("NewManualDialTableID"))
			//console.log(myIFrameOne.contentDocument.getElementById("vicidial_form");
			//var myIframe = document.getElementById('t_new');
			//myIframe.getElementById('MDPhonENumbeR').value=944267001;
			//myIframe.document.getElementById("MDDiaLCodE").value="51";
			//console.log(myIframe);
		}
		
		function getValuesIframe(){
			
		}
		
		</script>
		
	</body>
</html>