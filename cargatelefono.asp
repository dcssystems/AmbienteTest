<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("cargatelefono.asp") then
	buscador=obtener("buscador")	
	
	sql="select a.usuario, a.administrador, b.codagencia,b.razonsocial from usuario a left join agencia b on a.codagencia = b.codagencia where codusuario = "& session("codusuario")
	consultar sql,RS	
	usuario=rs.fields("usuario")
	fadmin=rs.fields("administrador")
	''agencia=iif(isNull(rs.Fields("razonsocial")),"Sin asignar",rs.Fields("razonsocial"))
	''codagencia=rs.fields("codagencia")
	rs.close
	
		if obtener("agregardato")<>"" then
		archivotxt=obtener("archivotxt")
			if archivotxt = "OK" then
				alerta = "<font face=Arial size=2 color=#339900>Se carg� el archivo correctamente a nuestro servidor.<br>Al finalizar la actualizaci�n de la Base de Datos, se enviar� un email con el estado final.</font>"
			else
				alerta ="<font face=Arial size=2 color=#CC0000>Ocurri� un error en la carga del archivo.<br>Int�ntelo nuevamente.</Font>"
			end if				
		end if
		''else
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
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
		
		

			<title>Nueva Carga de telefonos</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
			    function agregar() {
			        if (checkFile() == true) {
			            document.formula.agregardato.value = 1;
			            document.formula.submit();
			        }
			    }
			    function trim(string) {
			        while (string.substr(0, 1) == " ")
			            string = string.substring(1, string.length);
			        while (string.substr(string.length - 1, 1) == " ")
			            string = string.substring(0, string.length - 2);
			        return string;
			    }
			    function isEmailAddress(theElement) {
			        var s = theElement.value;
			        var filter = /^[\w-\.]+@([\w-]+\.)+[\w-]{2,4}$/ig;
			        if (s.length == 0) return true;
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
			<script language="javascript">
			    function checkFile() {
			        var fileElement = document.getElementById("archivotxt");
			        var fileExtension = "";
			        if (fileElement.value.lastIndexOf(".") > 0) {
			            fileExtension = fileElement.value.substring(fileElement.value.lastIndexOf(".") + 1, fileElement.value.length);
			        }
			        if (fileExtension == "txt" || fileExtension == "TXT") {
			            return true;
			        }
			        else {
			            alert("Debe seleccionar un archivo con extensi�n *.txt");
			            return false;
			        }
			    }
			</script>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
					<table border=0 cellspacing=0 cellpadding=0 width=100% height=30%>
					<form name=formula method=post enctype="multipart/form-data" action="uploadfiletelefono.asp">
					<tr>	
						<td bgcolor="#F5F5F5" colspan=3>			
							<font size=2 color=#483d8b face=Arial><b>&nbsp;Carga de Tel�fonos</b></font>
						</td>
					</tr>
					<!--<tr>
						<td width=15%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Usuario:</font></td>
						<td width=25%><font face=Arial size=2 color=#483d8b>&nbsp;<%=usuario%></font></td>
						<td rowspan=4 valign=bottom><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a></td>
					</tr>-->
					
					<tr>
						<td width=15%><font face=Arial size=2 color=#483d8b>&nbsp;&nbsp;Seleccione archivo:</font></td>
						<td width=25%><input type="file" name="archivotxt" id="archivotxt" accept=".txt"><font face=Arial size=2 color=#483d8b>&nbsp;(*.txt)</font></td>
					
						<td valign=bottom><a href="javascript:agregar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar"></a></td>
					</tr>										
					<tr>					
						<td colspan=3 bgcolor="#F5F5F5"><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>								
					</tr>
					<%if alerta<>"" then%>
					<tr>			
						<td colspan=3 align=center><b><%=alerta%></b></td>								
					</tr>
					<%end if%>						
						<input type=hidden name="agregardato" value="">
						<input type=hidden name="codtelefono" value="<%=codtelefono%>">
						<input type=hidden name="vistapadre" value="<%=obtener("vistapadre")%>">
						<input type=hidden name="paginapadre" value="<%=obtener("paginapadre")%>">
					</form>						
					</table>
			</body>
		</html>	
		<%		
		''end if
	else
	%>
	<script language="javascript">
	    alert("Ud. No tiene autorizaci�n para este proceso.");
	    window.open("userexpira.asp", "_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
    alert("Tiempo Expirado");
    //window.open("index.html","_top");
    window.open("index.html", "sistema");
    window.close();
</script>
<%
end if
%>

