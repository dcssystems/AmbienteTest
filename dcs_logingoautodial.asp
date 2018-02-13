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
	
		<form name="vicidial_form" id="vicidial_form" action="http://192.168.1.5/agent/agent.php" method="post" target="t_new">
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
			
			<!--MDDiaLCodE-->
			<!--MDPhonENumbeR-->
		
		</form>
		
		
		<table class="table table-striped" name="listCalling">
			<thead>
				<th>ITEM</th>
				<th>DNI</th>
				<th>NOMBRES</th>
				<th>APELLIDOS</th>
				<th>DEUDA</th>
				<th>TELEFONO</th>
				<th>ACCION</th>
			</thead>
			<tbody>
				<tr>
					<td>1</td>
					<td>45612379</td>
					<td>GINO</td>
					<td>VERA</td>
					<td>5200</td>
					<td class="phoneNumber1">987456321</td>
					<td>
						<input type="button" value="llamar" onclick="llamar('.phoneNumber1');">
						<input type="button" value="code" onclick="codeCountry();">
					</td>
				</tr>
				<tr>
					<td>2</td>
					<td>45698712</td>
					<td>MOISES</td>
					<td>MIMBELA</td>
					<td>5000</td>
					<td class="phoneNumber2">963153879</td>
					<td><input type="button" value="llamar" onclick="llamar('.phoneNumber2');"></td>
				</tr>
				<tr>
					<td>3</td>
					<td>43287388</td>
					<td>LUIS</td>
					<td>RUIZ</td>
					<td>4500</td>
					<td class="phoneNumber3">944267001</td>
					<td><input type="button" value="llamar" onclick="llamar('.phoneNumber3');"></td>
				</tr>
				<tr>
					<td>4</td>
					<td>47852369</td>
					<td>VICTOR</td>
					<td>TIBURCIO</td>
					<td>6400</td>
					<td class="phoneNumber4">987456321</td>
					<td><input  type="button" value="llamar" onclick="llamar('.phoneNumber4');"></td>
				</tr>
				
			</tbody>
		</table>
		
		
		<input type="hidden" name="MDPhonENumbeRHiddeN" id="MDPhonENumbeRHiddeN" value="" />
		<input type="hidden" name="MDLeadID" id="MDLeadID" value="" />
		<input type="hidden" name="MDType" id="MDType" value="" />
		
		
		
		<iframe name="t_new" style="height:100%; width:100%;" ></iframe><!--style="visibility:hidden;"-->
		
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
			var myIframe = window.parent.document.getElementById("MDDiaLCodE").value="51";
			//myIframe.document.getElementById("MDDiaLCodE").value="51";
			console.log(myIframe);
		}
		</script>
		
	</body>
</html>