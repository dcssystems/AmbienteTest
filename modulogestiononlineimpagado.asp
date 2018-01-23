<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") then
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
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
		var nuevodir;
		var nuevotelf;
		var nuevogestion;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanamarcas=global_popup_IWTSystem(ventanamarcas,"nuevomarcas.asp?vistapadre=" + window.name + "&paginapadre=admmarcas.asp&codmarca=" + codigo,"Newmarca","scrollbars=no,scrolling=no,top=" + ((screen.height - 140)/2 - 30) + ",height=140,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregardir()
		{
			nuevodir=global_popup_IWTSystem(nuevodir,"adicionardireccion.asp?vistapadre=" + window.name + "&paginapadre=modulogestiononlineimpagado.asp","Newdir","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
		}
		function agregartelf()
		{
			nuevotelf=global_popup_IWTSystem(nuevotelf,"adicionartelefono.asp?vistapadre=" + window.name + "&paginapadre=modulogestiononlineimpagado.asp","Newtelf","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
		}
		function agregargestion()
		{
			nuevogestion=global_popup_IWTSystem(nuevogestion,"adicionarseguimiento.asp?vistapadre=" + window.name + "&paginapadre=modulogestiononlineimpagado.asp","Newgestion","scrollbars=no,scrolling=no,top=" + ((screen.height - 140)/2 - 30) + ",height=140,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}	
		function actualizar()
		{
			document.formula.actualizarlista.value=1;
			document.formula.submit();
		}	
		function exportar()
		{
			document.formula.expimp.value=1;
			document.formula.submit();
		}	
		function imprimir()
		{
			window.open("impusuarios.asp","ImpUsuarios","scrollbars=yes,scrolling=yes,top=0,height=200,width=200,left=0,resizable=yes");
		}					
		function buscar()
		{
			document.formula.pag.value=1;
			document.formula.submit();
		}		
		function filtrar()
		{
			if (filtrardatos[0]==1)
			{
				filtrardatos[0]=0;
				dibujarTabla(0);
			}
			else
			{
				filtrardatos[0]=1;
				dibujarTabla(0);
			}
		}				
		function mostrarpag(pagina)
		{
			document.formula.pag.value=pagina;
			document.formula.submit();
		}
		</script>
		<style>
		A {
			FONT-SIZE: 12px; COLOR: #00529B; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
		}
		A:visited {
			TEXT-DECORATION: none; COLOR: #00529B;
		}
		A:hover {
			COLOR: #00529B; FONT-FACE:"Arial"; TEXT-DECORATION: none
		}
		.skin {
		position:absolute;
		top: 0px;
		font-color:#FFFFFF;
		font-size:12px;
		width:78%;
		height:50%;
		border:0px none #667ec5;
		text-align:left;
		font-family:Arial;
		line-height:16px;
		cursor:hand;
		visibility:hidden;
		background:#FFFFFF;
		}
		TABLE
		{
			border-width: 0px;
			border-style: none;
		}
		TH 
		{
			color:#FFFFFF;
			background: #007DC5;
			font-size:12px;
			font-family:Arial;
			cursor:hand;
		}
		
		</style>
		</head>
				
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0 border=1>		
			<tr bgcolor="#007DC5">
				<td colspan="9" align="center" height="18"><font size=2 face=Arial color="#FFFFFF"><b>Módulo Gestión On-Line Impagado</b></font>
				</td>
			</tr>
			<tr bgcolor="#BEE8FB">
				 <td align="right"><font size=2 face=Arial color=#00529B>Fecha&nbsp;de&nbsp;Proceso:</font></td>
				 <td><font size=2 face=Arial color=#00529B><b>14/08/2014</b></font></td>
				  <td>&nbsp;</td>
				    <td>&nbsp;</td>
				    <td>&nbsp;</td>
				 <td align="right"><font size=2 face=Arial color=#00529B>Email:</font></td>
				 <td><font size=2 face=Arial color=#00529B><b>ABC@DGH.COM</b></font></td>
				  <td align="right"><font size=2 face=Arial color=#00529B>Asignación</font></td>
				   <td><font size=2 face=Arial color=#00529B><b>REDITOS</b></font></td>
			</tr>
			<tr bgcolor="#BEE8FB">	
				 <td align="right"><font size=2 face=Arial color=#00529B>Territorio:</font></td>
				 <td><font size=2 face=Arial color=#00529B><b>SUR</b></font></td>
				  <td align="right"><font size=2 face=Arial color=#00529B>Oficina&nbsp;Gestora:</font></td>
				    <td><font size=2 face=Arial color=#00529B><b>Of. Cuzco</b></font></td>
				    <td>&nbsp;</td>
				 <td align="right"><font size=2 face=Arial color=#00529B>Grupo&nbsp;de&nbsp;riesgo de Cobranza:</font></td>
				 <td><font size=2 face=Arial color=#00529B><b>GRUPO 01</b></font></td>
				  <td align="right"><font size=2 face=Arial color=#00529B>Acciónx</font></td>
				   <td><font size=2 face=Arial color=#00529B><b>LLAMADA MAS VISITA<b></font></td>
			</tr>
			<tr bgcolor="#BEE8FB">	
				 <td align="right"><font size=2 face=Arial color=#00529B>Nombres:</font></td>
				 <td colspan=3><font size=2 face=Arial color=#00529B><b>A. RAMIRO OLIVERA FERNANDEZ BACA</b></font></td>
				 <td>&nbsp;</td>
				  <td align="right"><font size=2 face=Arial color=#00529B>Tipo&nbsp;documento:</font></td>
				  <td><font size=2 face=Arial color=#00529B><b>L</b></font></td>
				  <td align="right"><font size=2 face=Arial color=#00529B>N° documento:</font></td>
				   <td><font size=2 face=Arial color=#00529B><b>2383330<b></font></td>
			</tr>
			<tr bgcolor="#BEE8FB">	
				 <td align="right"><font size=2 face=Arial color=#00529B>Dirección:</font></td>
				 <td colspan="8"><select name="direccion" style="font-size: x-small; width: 85%;">
				<option value="">dir 1</option>
				</select>&nbsp;<input type =button name=nuevodir value='nuevodir' onclick="javascript:agregardir();"></font></td>
			</tr>
			<tr bgcolor="#BEE8FB">
			 <td align="right"><font size=2 face=Arial color=#00529B>Teléfono:</font></td>
			<td colspan="8"><select name="direccion" style="font-size: x-small; width: 85%;">
				<option value="">particular			983736475</option>
				<option value="">privado			43534534</option>
				<option value="">telf 1</option>
				<option value="">telf 1</option>
				<option value="">telf 1</option>
				<option value="">telf 1</option>
				<option value="">telf 1</option>
				<option value="">telf 1</option>
				<option value="">telf 1</option>
				</select>
				<input type =button name=nuevotelf value='nuevotelf' onclick="javascript:agregartelf();"></font></td>
			<tr>
			</table>
			<table width=100% cellpadding=4 cellspacing=0 border=1>
			<tr bgcolor="#007DC5">
				<td colspan="13" align="left" height="18"><font size=2 face=Arial color="#FFFFFF"><b>Obligaciones del Cliente</b></font>
				</td>
			<tr bgcolor="#007DC5">
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Cuotas</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">N° Contrato</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Producto</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Primer venc.</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Días venc</font></td>				  
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Último venc.</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Capital</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Interés</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Gastos&nbsp;y Comis.</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Int.&nbsp;Comp.&nbsp;y Mora:</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Moneda</font></td>
					<td align="right"><font size=2 face=Arial color="#FFFFFF">Total</font></td>				  
					<td align="right"><font size=2 face=Arial color="#FFFFFF">&nbsp;</font></td>
					</tr>
					<tr bgcolor="#BEE8FB">
					<td align="right"><font size=2 face=Arial color=#00529B>Cuotas</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>N° Contrato</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Producto</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Primer venc.</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Días venc</font></td>				  
					<td align="right"><font size=2 face=Arial color=#00529B>Último venc.</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Capital</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Interés</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Gastos&nbsp;y Comis.</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Int.&nbsp;Comp.&nbsp;y Mora:</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Moneda</font></td>
					<td align="right"><font size=2 face=Arial color=#00529B>Total</font></td>				  
					<td align="right"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
					</tr>
			</table>
			<table width=100% cellpadding=4 cellspacing=0 border=1>
			<tr bgcolor="#007DC5">
				<td colspan="13" align="left" height="18"><font size=2 face=Arial color="#FFFFFF"><b>Historial de Gestiones</b></font>
				</td>
			<tr bgcolor="#BEE8FB">
			<td align="right"><font size=2 face=Arial color=#00529B><b>Fecha inicio:</td>
			<td width="78"><input name="fechapromesaini" type=text maxlength=10 id="sel1"  value="<%if IsDate(fechapromesaini) then%><%=fechapromesaini%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel1', '%d/%m/%Y');">
			<td align="right"><font size=2 face=Arial color=#00529B><b>Fecha fin:</td>
			<td width="78"><input name="fechapromesaini" type=text maxlength=10 id="sel2"  value="<%if IsDate(fechapromesaini) then%><%=fechapromesaini%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel2', '%d/%m/%Y');">
			<td align="right">&nbsp;&nbsp;<input type =button name=nuevogestion value='nuevogestion' onclick="javascript:agregargestion();"></td>
			<tr>
			</table>
		</form>
		<script language="javascript">
			inicio();
		</script>
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>							
		</body>
		</html>		
		<%
		''Codigo exp excel
		''Si se pide exportar a excel
				if expimp="1" then

					''se coloca arriba para el enlace si no abre directo el archivo
					''sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar' or descripcion='RutaWebExportar'"
					''consultar sql,RS
					''RS.Filter=" descripcion='RutaFisicaExportar'"
					''RutaFisicaExportar=RS.Fields(1)
					''RS.Filter=" descripcion='RutaWebExportar'"
					''RutaWebExportar=RS.Fields(1)	
					''RS.Filter=""			
					''RS.Close					
					''Para Exportar a Excel
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'Cod.Marca','Descripción' "
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select codmarca,descripcion " & _
								 "from CobranzaCM.dbo.marca " & filtrobuscador & " order by codmarca" 
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					conn.execute sql

					''Tercero borrar UserExport*.xls
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
										
					''Cuarto Uno los 2 archivos en temp*.txt
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''copy " & chr(34) & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt" & chr(34) & " + " & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & " /b''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Quinto Elimino los 2 archivos en temp*.txt
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql										
				%>
					<script language="javascript">
						window.open("<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>","_self");
					</script>
				<%					
				end if			
		%>		
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
	window.open("index.html","_top");
</script>
<%
end if
%>



