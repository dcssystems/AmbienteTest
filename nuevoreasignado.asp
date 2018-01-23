<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admreasignado.asp") then
	buscador=obtener("buscador")	
	codreasignado=obtener("codreasignado")
		if obtener("agregardato")<>"" then
		codagencia=obtener("codagencia")
		codigocentral=obtener("codigocentral")
		fechacaduca=obtener("fechacaduca")
		segmento_gestion=obtener("segmento_gestion")
		
		if segmento_gestion<>"" then
		    xsegmento_gestion="'" & trim(segmento_gestion) & "'"
		else
		    xsegmento_gestion="null"
		end if
		
		
		
		if obtener("activo")<>"" then activo="1" else activo="0" end if	
						
			existereasignado=0
			
			if codreasignado="" then
				sql="select count(*) from reasignado where codigocentral='" & codigocentral & "'"
			else
				sql="select count(*) from reasignado where codigocentral='" & codigocentral & "' and codreasignado<>" & codreasignado
			end if			
			consultar sql,RS
			existereasignado=RS.Fields(0)
			RS.Close			
			if existereasignado=0 then			
				if obtener("agregardato")="1" then	
					sql="insert into reasignado (codagencia,codigocentral,fechacaduca,usuarioregistra,fecharegistra,activo,segmento_gestion) values (" & codagencia & ",'" & codigocentral & "','" & mid(obtener("fechacaduca"),7,4) & mid(obtener("fechacaduca"),4,2) & mid(obtener("fechacaduca"),1,2) & "'," & session("codusuario") & ",getdate()," & activo & "," & xsegmento_gestion & ")"
				else
					sql="update reasignado set codagencia=" & codagencia & ",codigocentral='" & codigocentral & "',fechacaduca='" & mid(obtener("fechacaduca"),7,4) & mid(obtener("fechacaduca"),4,2) & mid(obtener("fechacaduca"),1,2) & "',usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate(), activo=" & activo & ", segmento_gestion=" & xsegmento_gestion & " where codreasignado=" & codreasignado
				end if
				''Response.Write sql
				conn.execute sql
									
				%>
				<script language=javascript>
					<%if obtener("agregardato")="1" then%>
					//alert("Se agregó el usuario correctamente.");
					<%else%>
					//alert("Se modificó el usuario correctamente.");
					<%end if%>				
					<%if obtener("paginapadre")="admreasignado.asp" then%>window.open("<%=obtener("paginapadre")%>","<%=obtener("vistapadre")%>");<%end if%>
					window.close();
				</script>			
				<%
			else
			%>
				<script language=javascript>
					alert("El reasignado ya existe.");
					history.back();
				</script>			
			<%				
			end if
		else
			if codreasignado<>"" then
					sql="select A.*,B.nombres as Nombreusureg, B.apepaterno as Apepatusureg, B.apematerno as Apematusureg, C.nombres as Nombreusumod,C.apepaterno as Apepatusumod, C.apematerno as Apematusumod from reasignado A inner join usuario B on B.codusuario=A.usuarioregistra left outer join usuario C on C.codusuario=A.usuariomodifica where A.codreasignado = " & codreasignado
					consultar sql,RS
					codagencia=rs.Fields("codagencia")		
					codigocentral=rs.Fields("codigocentral")
					segmento_gestion=rs.Fields("segmento_gestion")
					fechacaduca=rs.Fields("fechacaduca")
					fechaReg=RS.Fields("fecharegistra")
					usuarioReg=iif(IsNull(RS.Fields("Nombreusureg")),"",RS.Fields("Nombreusureg")) & ", " & iif(IsNull(RS.Fields("Apepatusureg")),"",RS.Fields("Apepatusureg")) & " " & iif(IsNull(RS.Fields("Apematusureg")),"",RS.Fields("Apematusureg"))
					fechaMod=RS.Fields("fechamodifica")
					usuarioMod=iif(IsNull(RS.Fields("Nombreusumod")),"",RS.Fields("Nombreusumod")) & ", " & iif(IsNull(RS.Fields("Apepatusumod")),"",RS.Fields("Apepatusumod")) & " " & iif(IsNull(RS.Fields("Apematusumod")),"",RS.Fields("Apematusumod"))
					activo=rs.Fields("activo")
					RS.Close
			else
				activo="1"
			end if		
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title><%if codreasignado="" then%>Nueva <%end if%>reasignado</title>
			<link rel="stylesheet" type="text/css" media="all" href="scripts/calendar-blue2.css" title="blue" />
		<script type="text/javascript" src="scripts/calendar.js"></script>
		<script type="text/javascript" src="scripts/calendar-es.js"></script>
		<script type="text/javascript">
		    function selected(cal, date) {
		        cal.sel.value = date;
		        if (cal.dateClicked && (cal.sel.id == "sel1" || cal.sel.id == "sel3"))
		            cal.callCloseHandler();
		    }
		    function closeHandler(cal) {
		        cal.hide();
		        _dynarch_popupCalendar = null;
		    }
		    function showCalendar(id, format, showsTime, showsOtherMonths) {
		        var el = document.getElementById(id);
		        if (_dynarch_popupCalendar != null) {
		            _dynarch_popupCalendar.hide();
		        } else {
		            var cal = new Calendar(1, null, selected, closeHandler);
		            if (typeof showsTime == "string") {
		                cal.showsTime = true;
		                cal.time24 = (showsTime == "24");
		            }
		            if (showsOtherMonths) {
		                cal.showsOtherMonths = true;
		            }
		            _dynarch_popupCalendar = cal;
		            cal.setRange(1900, 2070);
		            cal.create();
		        }
		        _dynarch_popupCalendar.setDateFormat(format);
		        _dynarch_popupCalendar.parseDate(el.value);
		        _dynarch_popupCalendar.sel = el;

		        _dynarch_popupCalendar.showAtElement(el.nextSibling, "Br");
		        return false;
		    }
		</script>
			<script language=javascript>
				var limpioclave=0;
				<%if codreasignado="" then%>
				function agregar()
				{
					
					if(trim(formula.codigocentral.value)==""){alert("Debe ingresar un Cliente para ingreso al Sistema.");return;}													
					if(trim(formula.fechacaduca.value)==""){alert("Debe ingresar una Fecha para ingreso al Sistema.");return;}	
					
					document.formula.agregardato.value=1;
					document.formula.submit();
				}
				<%else%>
				function modificar()
				{
					if(trim(formula.codigocentral.value)==""){alert("Debe ingresar un Cliente para ingreso al Sistema.");return;}													
					if(trim(formula.fechacaduca.value)==""){alert("Debe ingresar una Fecha para ingreso al Sistema.");return;}	
					
					document.formula.agregardato.value=2;
					document.formula.submit();
				}				
				<%end if%>
				function trim(string)
				{
					while(string.substr(0,1)==" ")
					string = string.substring(1,string.length) ;
					while(string.substr(string.length-1,1)==" ")
					string = string.substring(0,string.length-2) ;
					return string;
				}			
				function isEmailAddress(theElement)
				{
				var s = theElement.value;
				var filter=/^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ig;
				if (s.length == 0 ) return true;
				   if (filter.test(s))
				      return true;
				   else
					theElement.focus();
					return false;
				}					
			</script>
			<style>
			A {
				FONT-SIZE: 12px; COLOR: #483d8b; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
			}
			A:visited {
				TEXT-DECORATION: none; COLOR: #483d8b;
			}
			A:hover {
				COLOR: #483d8b; FONT-FACE:"Arial"; TEXT-DECORATION: none
			}			
			</style>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=100%>
					<form name=formula method=post action="nuevoreasignado.asp">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=2>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;<b><%if codreasignado="" then%>Nuevo <%end if%>Reasignado</b></b></font>
						</td>
					</tr>
					<%if fechaReg<>"" then%>
					<tr height=20>
						<td colspan=2 align=right><font face=Arial size=1 color=#483d8b>Registró:&nbsp;<b><%=usuarioReg%>&nbsp;el&nbsp;<%=fechaReg%></b>
						<%if fechaMod<>"" then%><BR>Modificó:&nbsp;<b><%=usuarioMod%>&nbsp;el&nbsp;<%=fechaMod%></b><%end if%>
						</font></td>
					</tr>	
					<%end if%>						
					<%if codreasignado = "" then%>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código:</font></td>
						<td>
							<font face=Arial size=2 color=#483d8b>Nuevo</font></td>
					</tr>
					
					<%else%>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código:</font></td>
						<td>
							  
							  <input type=hidden name="codreasignado" value="<%=codreasignado%>">
							  <font face=Arial size=2 color=#483d8b><%=codreasignado%></font></td>
						
					</tr>
					<%end if%>
					<tr>
					<td bgcolor ="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Agencia:</font></td>
					<td bgcolor ="#f5f5f5">
						<select name = "codagencia" style="font-size: xx-small; width: 200px;">
						<%
						sql = "select codagencia, razonsocial from agencia order by codagencia"
						consultar sql,RS
						Do While Not  RS.EOF
						%>
						<option value="<%=RS.Fields("codagencia")%>" <% if codagencia<>"" then%><% if RS.fields("codagencia")=int(codagencia) then%> selected<%end if%><%end if%>><%=RS.Fields("razonsocial")%></option>
						<%
						RS.MoveNext
						loop
						RS.Close
						%>
						</select>
						</td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Código Central:</font></td>
						<td><input name="codigocentral" type=text maxlength=200 value="<%=codigocentral%>" style="font-size: xx-small; width: 200px;"></td>
					</tr>
					<tr>
						<td bgcolor ="#f5f5f5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Caduca:</font></td>
						<td bgcolor ="#f5f5f5"><input name="fechacaduca" id="sel0" readonly  type=text maxlength=10 size=10 value="<%=fechacaduca%>"  style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel0', '%d/%m/%Y');" id=image1 name=image1></td>
					</tr>
					<tr>
						<td width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Segmento Gestión:</font></td>
						<td><input name="segmento_gestion" id="Text1" readonly  type=text maxlength=10 size=10 value="<%=segmento_gestion%>"  style="font-size: x-small; width: 60px;"></td>
					</tr>
					<tr>
						<td bgcolor="#F5F5F5" width=30%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Estado:</font></td>
						<td bgcolor="#F5F5F5"><input type=checkbox name="activo" style="font-size: xx-small;" <%if activo=1 then%> checked<%end if%>>&nbsp;&nbsp;<font face=Arial size=2 color=#483d8b>Activo</font></td>
					</tr>			
					<tr>					
						<td><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
						<td align=right height=40><%if codreasignado="" then%><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%else%><a href="javascript:modificar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a>&nbsp;<%end if%><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>					
					</tr>
						<input type=hidden name="agregardato" value="">
						<input type=hidden name="vistapadre" value="<%=obtener("vistapadre")%>">
						<input type=hidden name="paginapadre" value="<%=obtener("paginapadre")%>">
					</form>						
					</table>
					
			</body>
		</html>	
		<%		
		end if
	else
	%>
	<script language="javascript">
		alert("Ud. No tiene autorización para este proceso.");
		window.open("userexpira.asp","_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	//window.open("index.html","_top");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>

