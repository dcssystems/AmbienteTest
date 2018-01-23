<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") then
		codigocentral=obtener("codigocentral")
		codrespgestion=obtener("codrespgestion")
		contrato=obtener("contrato")
		fechadatos=obtener("fechadatos")
		fechagestion=obtener("fechagestion")
		codtipoaccion=obtener("codtipoaccion")
		codtipocontacto=obtener("codtipocontacto")
		codgestion=obtener("codgestion")
		
		sql1="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload' or descripcion='RutaWebUpload'"
		consultar sql1,RS3
		RS3.Filter=" descripcion='RutaFisicaUpload'"
		RutaFisicaUpload=RS3.Fields(1)
		RS3.Filter=" descripcion='RutaWebUpload'"
		RutaWebUpload=RS3.Fields(1)				
		RS3.Filter=""
		RS3.Close	
		
		sql="select A.codrespgestion,D.descripcion AS tipoaccion,C.descripcion AS tipocontacto,(CASE WHEN ltrim(rtrim(A.fono))<>'' THEN rtrim(ltrim(A.tipofono)) + '-' + rtrim(ltrim(A.prefijo)) + '-' + rtrim(ltrim(A.fono)) + (CASE WHEN LEN(RTRIM(ltrim(A.extension)))> 0 and A.extension<>'0000' THEN '-'+ RTRIM(LTRIM(A.extension)) ELSE '' END) ELSE A.direccion + ' - ' + UB.distrito + ' - ' + UB.provincia + ' - ' + UB.departamento END) AS dirtef,B.descripcion AS respgestion,A.comentario as comentario,A.ficherogestion AS ficherogestion from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion inner join TipoContacto C on B.codtipocontacto=C.codtipocontacto inner join TipoAccion D on B.codtipoaccion=D.codtipoaccion LEFT OUTER JOIN Ubigeo AS UB ON UB.coddpto = A.coddpto AND UB.codprov = A.codprov AND UB.coddist = A.coddist where A.codrespgestion=" & codrespgestion
		consultar sql,RS
		if not RS.EOF then
		    tipoaccion=RS.Fields("tipoaccion")
		    tipocontacto=RS.Fields("tipocontacto")
		    dirtef=RS.Fields("dirtef")
		    respgestion=RS.Fields("respgestion")
		    comentario=RS.Fields("comentario")
		    ficherogestion=RS.Fields("ficherogestion")
		end if
		RS.close
		%>
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>Gestión</title>
			<script language='javascript' src="scripts/popcalendar.js"></script> 
			<script language=javascript>
							
				
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
			
						
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
				<form name=formula method=post action="verseguimiento.asp">
				    <table width=100% height=80% border=0 cellspacing=0 cellpadding=0>
				        <tr bgcolor="#007DC5">
					        <td align="middle" height="22" colspan=2><font size=2 face=Arial color="#FFFFFF"><b>Gestión</b></font></td>
					    </tr>
					    <tr>
				            <td width=22%><font face=Arial size=2 color=#483d8b>&nbsp;Acción realizada:</font></td>
				            <td><font face=Arial size=2 color=#483d8b><b><%=tipoaccion%></b></font></td>
			            </tr>
			            <tr>
				            <td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;Tipo de contacto:</font></td>
				            <td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b><b><%=tipocontacto%></b></font></td>
			            </tr>
			            <tr>
				            <td width=22%><font face=Arial size=2 color=#483d8b>&nbsp;Direccion/Telefono:</font></td>
				            <td><font face=Arial size=2 color=#483d8b><b><%=dirtef%></b></font></td>
			            </tr>
			            <tr>
				            <td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b>&nbsp;Respuesta Gestion:</font></td>
				            <td bgcolor="#f5f5f5"><font face=Arial size=2 color=#483d8b><b><%=respgestion%></b></font></td>
				            
			            </tr>
			            <tr>
				            <td width=22%><font face=Arial size=2 color=#483d8b>&nbsp;Comentario:</font></td>
				            <td><font face=Arial size=2 color=#483d8b><b><%=comentario%></b></font></td>
				            
			            </tr>
			            <%if not IsNull(ficherogestion) then%>
			                <tr>
				                <td align=center valign=top><font face=Arial size=2 color=#483d8b>&nbsp;</font></td>
				                <td align=left><font face=Arial size=2 color=#483d8b><b>Archivo adjunto:&nbsp;<a href="<%=RutaWebUpload%>/<%=ficherogestion%>" target="T_New"><img src="imagenes/descargarpeq.png" border=0 alt="Descargar Archivo" title="Descargar Archivo"></a></b></font></td>
			                 </tr>
			            <%end if%>
			        </table>
			        <table width=100% height=20% border=0 cellspacing=0 cellpadding=0>
			        <tr>					
						    <td bgcolor="#f5f5f5" align=right><a href="javascript:window.close();"><img src="imagenes/salida.gif" border=0 alt="Salir" title="Salir"></a>&nbsp;</td>
					</tr>
			        </table>
			</body>
		</html>	
		<%		
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


