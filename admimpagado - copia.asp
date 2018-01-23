<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") then
		contrato=obtener("contrato")
		codigocentral=obtener("codigocentral")
		codgestor=obtener("codgestor")
		codtipodocumento=obtener("codtipodocumento")
		numdocumento=obtener("numdocumento")
		diaatrasoini=obtener("diaatrasoini")
		diaatrasofin=obtener("diaatrasofin")
		codproducto=obtener("codproducto")
		codagencia=obtener("codagencia")
		codmarca=obtener("codmarca")
		codterritorio=obtener("codterritorio")
		if codterritorio="" then
			codoficina=""
		else
			codoficina=obtener("codoficina")
		end if
		tipogestion=obtener("tipogestion")
		fechaasigini=obtener("fechaasigini")
		fechaasigfin=obtener("fechaasigfin")
		fechapromesaini=obtener("fechapromesaini")
		fechapromesafin=obtener("fechapromesafin")
		fechagestionini=obtener("fechagestionini")
		fechagestionfin=obtener("fechagestionfin")
		codgestion=obtener("codgestion")
%>
<html>
<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
<head>
<title>Seguimiento de Impagado</title>
 
<script language=javascript>
	function actualizaterritorio()
	{
		document.formula.submit();
	}
</script>
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






</head>
<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
<form name=formula method=post action="admimpagado.asp">
<table border=0 cellspacing=0 cellpadding=4 width=100%>
  <tr bgcolor="#007DC5">
    <td colspan="17" align="center" height="18">
      <font size=2 face=Arial color="#FFFFFF"><b>Gestión de Impagados</b></font>
    </td>
  </tr>
  <tr bgcolor="#BEE8FB">
    <td align="right"><font size=2 face=Arial color=#00529B><b>N°&nbsp;Contrato:</b></font></td>
    <td colspan="4"><input name="contrato" type=text maxlength=18 value="<%=contrato%>" style="font-size: x-small; width: 250px;"></td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Código&nbsp;Cliente:</b></font></td>
    <td colspan="4"><input name="codigocentral" type=text maxlength=8 value="<%=codigocentral%>" style="font-size: x-small; width: 250px;"></td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Gestor:</b></td>
    <td colspan="4"><select name="codgestor" style="font-size: x-small; width: 250px;"></td>
  </tr>
  <tr bgcolor="#BEE8FB">
    <td align="right"><font size=2 face=Arial color=#00529B><b>Tipo&nbsp;Documento:</b></font></td>
    <td colspan="4">
			<select name="codtipodocumento" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Tipo Documento</option>
				<%
				sql = "select codtipodocumento,descripcion from TipoDocumento where activo=1 order by descripcion"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
					<option value="<%=RS.Fields("codtipodocumento")%>" <% if codtipodocumento<>"" then%><% if RS.fields("codtipodocumento")=int(codtipodocumento) then%> selected<%end if%><%end if%>><%=RS.Fields("CodTipoDocumento") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
				%>				
			</select>
    </td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>N°&nbsp;Documento:</b></font></td>
    <td colspan="4"><input name="numdocumento" type=text maxlength=200 value="<%=nrocontrato%>" style="font-size: x-small; width: 250px;"></td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Días&nbsp;de&nbsp;Atraso:</b></font></td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Del</b></font></td>
  <td><input name="diaatrasoini" type=text maxlength=10 value="<%=diaatrasoini%>" style="text-align: right;font-size: x-small; width: 50px;"></td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>al</b></font></td>
    <td><input name="diaatrasofin" type=text maxlength=10 value="<%=diaatrasofin%>" style="text-align: right;font-size: x-small; width: 50px;"></td>
  </tr>
  <tr bgcolor="#BEE8FB">
    <td align="right"><font size=2 face=Arial color=#00529B><b>Producto</b></font></td>
    <td colspan="4">
    <select name="producto" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Producto</option>
				<%
				sql = "select codproducto, descripcion from producto where activo=1 order by codproducto"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
					<option value="<%=RS.Fields("codproducto")%>" <% if codproducto<>"" then%><% if RS.fields("codproducto")=codproducto then%> selected<%end if%><%end if%>><%=RS.Fields("codproducto") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
				%>
		</select>
	</td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Agencia:</b></font></td>
    <td colspan="4">
		<select name = "agencia" style="font-size: x-small; width: 250px;">
        <option value="">Seleccione Agencia</option>
   <%
    sql = "select codagencia, razonsocial from agencia where activo=1 order by razonsocial"
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
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Marca:</b></font></td>
    <td colspan="4">
			<select name="codmarca" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Marca</option>
				<%
				sql = "select codmarca,descripcion from Marca where activo=1 order by descripcion"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
					<option value="<%=RS.Fields("codmarca")%>" <% if codmarca<>"" then%><% if RS.fields("codmarca")=int(codmarca) then%> selected<%end if%><%end if%>><%=RS.Fields("codmarca") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
				%>				
			</select>    
    </td>
  </tr>
  <tr bgcolor="#BEE8FB">
    <td align="right"><font size=2 face=Arial color=#00529B><b>Territorio:</b></font></td>
    <td colspan="4">
		<select name="codterritorio" onchange="actualizaterritorio()" style="font-size: x-small; width: 250px;">
			<option value="">Seleccione Territorio</option>
			<%
				sql = "select codterritorio, descripcion from territorio where activo=1 order by codterritorio"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
				<option value="<%=RS.Fields("codterritorio")%>" <%if codterritorio<>"" then%><% if RS.fields("codterritorio")=codterritorio then%> selected<%end if%><%end if%>><%=RS.Fields("codterritorio") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
			%>
		</select>
    </td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Oficina:</b></font></td>
    <td colspan="4">
		<select name = "codoficina" style="font-size: x-small; width: 250px;">
		<option value="">Seleccione Oficina</option>
        <%
			if codterritorio<>"" then
				sql = "select codoficina, descripcion from oficina where activo = 1 and codterritorio = " & codterritorio & " order by codoficina"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
				<option value="<%=RS.Fields("codoficina")%>" <% if codoficina<>"" then%><% if RS.fields("codoficina")=int(codoficina) then%> selected<%end if%><%end if%>><%=RS.Fields("CodOficina") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
			end if
		%>
		</select>
    </td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Tipo&nbsp;de&nbsp;Gestión:</b></font></td>
    <td colspan="4">
		<select name = "tipogestion" style="font-size: x-small; width: 250px;">
			<option value="">Seleccione Gestión</option>
			<option value="Preventiva" <% if tipogestion<>"" then%><%if tipogestion="Preventiva" then%> selected<%end if%><%end if%>>Preventiva</option>
			<option value="Impagado" <% if tipogestion<>"" then%><%if tipogestion="Impagado" then%> selected<%end if%><%end if%>>Impagado</option>
		</select>    
    </td>
  </tr>
  <tr bgcolor="#BEE8FB">
    <td align="right"><font size=2 face=Arial color=#00529B><b>Fecha&nbsp;de&nbsp;Asignación:</b></font></td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Del</td>
    <td width="78"><input name="fechaasigini"  id="sel0" readonly  type=text maxlength=10 size=10 value="<%if IsDate(fechaasigini) then%><%=fechaasigini%><%else%><%=Date()%><%end if%>"  style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel0', '%d/%m/%Y');"></td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>al</td>
    <td>
		<input name="fechaasigfin" id="sel3" type=text maxlength=10 readonly value="<%if IsDate(fechaasigfin) then%><%=fechaasigfin%><%else%><%=Date()%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel3', '%d/%m/%Y');">
	<!--<input type="text" name="date3" id="sel3" size="10"><input type="image" value=" ... " onclick="return showCalendar('sel3', '%d/%m/%Y');">-->
	</td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Fecha&nbsp;de&nbsp;Promesa:</b></font></td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Del</td>
    <td width="78"><input name="fechapromesaini" type=text maxlength=10 id="sel4"  value="<%if IsDate(fechapromesaini) then%><%=fechapromesaini%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel4', '%d/%m/%Y');">
    </td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>al</td>
    <td><input name="fechapromesafin" type=text maxlength=10  id="sel5"  value="<%if IsDate(fechapromesafin) then%><%=fechapromesafin%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel5', '%d/%m/%Y');">
    </td>
    <td>&nbsp;</td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Fecha&nbsp;de&nbsp;Gestión:</b></font></td>
    <td align="right"><font size=2 face=Arial color=#00529B><b>Del</td>
    <td width="78"><input name="fechagestionini" type=text mask="00/00/0000" maxlength=10 id="sel6"  value="<%if IsDate(fechagestionini) then%><%=fechagestionini%><%end if%>"  style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel6', '%d/%m/%Y');">
    <td align="right"><font size=2 face=Arial color=#00529B><b>al</td>
    <td><input name="fechagestionfin" type=text maxlength=10 id="sel7"  value="<%if IsDate(fechagestionfin) then%><%=fechagestionfin%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel7', '%d/%m/%Y');">
  </tr>
  <tr bgcolor="#BEE8FB">
    <td align="right"><font size=2 face=Arial color=#00529B><b>Última&nbsp;Gestión:</b></font></td>
    <td colspan="11">
		<select name = "codgestion" style="font-size: x-small; width: 600px;">
			<option value="">Seleccione Última Gestión</option>
			<%
				sql = "select codgestion, descripcion from Gestion where activo=1 order by codgestion"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
				<option value="<%=RS.Fields("codgestion")%>" <% if codgestion<>"" then%><% if RS.fields("codgestion")=codgestion then%> selected<%end if%><%end if%>><%=RS.Fields("codgestion") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
			%>
		</select>       
    </td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
    <td>&nbsp;</td>
   </tr>
  <tr bgcolor="#F5F5F5">
    <td colspan="17" align="right" height="18"><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
  </tr>   
</table>
<input type=hidden name="solicitabuscar" value="">
</form>
</body>
<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>
</html>
<% end if
	desconectar
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	window.open("index.html","_top");
</script>
<%
end if
%>

