<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admida.asp") then
	buscaractivos=obtener("buscaractivos")
		buscador=obtener("buscador")
		expimp=obtener("expimp")
		
		if expimp="1" then
			sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar' or descripcion='RutaWebExportar'"
			consultar sql,RS
			RS.Filter=" descripcion='RutaFisicaExportar'"
			RutaFisicaExportar=RS.Fields(1)
			RS.Filter=" descripcion='RutaWebExportar'"
			RutaWebExportar=RS.Fields(1)				
			RS.Filter=""
			RS.Close	
			tiempoexport=Now()
		end if		
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<script language=javascript src="scripts/TablaDinamica.js"></script>
		<script type="text/javascript" src="scripts/tristate-0.9.2.js" ></script>
		<script language=javascript>
		var ventanaoficina;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanaoficina=global_popup_IWTSystem(ventanaoficina,"nuevooficina.asp?vistapadre=" + window.name + "&paginapadre=admoficina.asp&codoficina=" + codigo,"Newoficina","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=250,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregar()
		{
			ventanaoficina=global_popup_IWTSystem(ventanaoficina,"nuevooficina.asp?vistapadre=" + window.name + "&paginapadre=admoficina.asp","Newoficina","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=250,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
		TD
		{
			color:#00529B;
			font-size:12px;
			font-family:Arial;
		}
		TR
		{
			background: #FFFFFF;
		}
		</style>
		</head>
		
		<script language=javascript>
			rutaimgcab="imagenes/"; 
		  //Configuración general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=16;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('Territorio','Stock de cierre','Ingresos del mes','Meta mensual','Meta anual','Avance anual','Avance anual %','Avance mes','Avance mes %','REF','TRA','EJEC','EFE','OTRO','PART','PYME');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(true, true, true ,true,true,true,true,true,true,true,true,true,true,true,true,true);
		    anchocolumna[tabla] =  new Array( '5%', '5%', '5%' , '8%' ,'8%' ,'8%' ,'8%','3%' ,'3%','5%', '5%', '5%' , '8%' ,'8%' ,'8%','5%');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','center','left','left','left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('center','left','left','left','left','left','left','center','left','center','left','left','left','left','left','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','center','left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,2 ,2 ,0 , 0 , 2 , 2 , 2 , 2, 2 , 2 , 2 , 2, 2 , 2 , 2);
		    formatofecha[tabla] =   new Array(''  ,''   ,'' ,'','' ,'','','','',''  ,''   ,'' ,'','' ,'','');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][1]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][4]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][6]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][7]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][8]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][9]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][10]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][11]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][12]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][13]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][14]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][15]='<a href=javascript:modificar("-id-");>-valor-</a>';
					
					
		    filtrardatos[tabla]=0; //define si carga auto el filtro
		    filtrofomulario[tabla] = new Array();
		    tipofiltrofomulario[tabla] = new Array();
		    	filtrofomulario[tabla][0]='';
				filtrofomulario[tabla][1]='';
				filtrofomulario[tabla][2]=''; //objetofiltro("text",tabla,2,'contiene');
				filtrofomulario[tabla][3]='';
				filtrofomulario[tabla][4]='';
				filtrofomulario[tabla][5]='';
				filtrofomulario[tabla][6]='';
				filtrofomulario[tabla][7]='';
				filtrofomulario[tabla][8]='';
				filtrofomulario[tabla][9]=''; //objetofiltro("text",tabla,2,'contiene');
				filtrofomulario[tabla][10]='';
				filtrofomulario[tabla][11]='';
				filtrofomulario[tabla][12]='';
				filtrofomulario[tabla][13]='';
				filtrofomulario[tabla][14]='';
				filtrofomulario[tabla][15]='';
			
		    valorfiltrofomulario[tabla] = new Array();
				valorfiltrofomulario[tabla][0]='';
				valorfiltrofomulario[tabla][1]='';
				valorfiltrofomulario[tabla][2]='';
				valorfiltrofomulario[tabla][3]='';
				valorfiltrofomulario[tabla][4]='';
				valorfiltrofomulario[tabla][5]='';
				valorfiltrofomulario[tabla][6]='';
				valorfiltrofomulario[tabla][7]='';
				valorfiltrofomulario[tabla][8]='';
				valorfiltrofomulario[tabla][9]='';
				valorfiltrofomulario[tabla][10]='';
				valorfiltrofomulario[tabla][11]='';
				valorfiltrofomulario[tabla][12]='';
				valorfiltrofomulario[tabla][13]='';
				valorfiltrofomulario[tabla][14]='';
				valorfiltrofomulario[tabla][15]='';
							
		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
				
		
		contadortotal=0
		sql="select count(*) from Reporte_IDA A where A.fechadatos=(Select max(B.fechadatos) from Reporte_IDA B)"
		consultar sql,RS	
		contadortotal=rs.fields(0)
		RS.Close		
				
		cantidadxpagina=18
		paginasxbloque=10
		
		if obtener("pag")="" then
		pag=1
		else
		pag=int(obtener("pag"))
		end if
		
		topnovisible=int((pag - 1)*cantidadxpagina)
		
		if contadortotal mod cantidadxpagina = 0 then
		pagmax=int(contadortotal/cantidadxpagina)
		else
		pagmax=int(contadortotal/cantidadxpagina) + 1
		end if

		if contadortotal mod cantidadxpagina = 0 then
		pagmax=int(contadortotal/cantidadxpagina)
		else
		pagmax=int(contadortotal/cantidadxpagina) + 1
		end if

		if pag mod paginasxbloque = 0 then
		bloqueactual=int(pag/paginasxbloque)
		else
		bloqueactual=int(pag/paginasxbloque) + 1
		end if				

		if pagmax mod paginasxbloque = 0 then
		bloquemax=int(pagmax/paginasxbloque)
		else
		bloquemax=int(pagmax/paginasxbloque) + 1
		end if
		
		if pag>1 then					
		sql="select top " & cantidadxpagina & "  A.TERRITORIO, A.[Stock de cierre],A.[Ingresos del Mes],A.[Meta Mensual],A.[Meta Anual],A.[Avance Anual],A.[Avance Anual %],A.[Avance Mes],A.[Avance Mes %],A.REF,A.TRA,A.EJEC,A.EFE,A.OTRO,A.PART,A.PYME FROM Reporte_IDA A where A.FECHADATOS=(Select max(B.fechadatos) from Reporte_Ida B) and A.Territorio not in (select top " & topnovisible & " A1.Territorio from REPORTE_IDA A1 where A1.fechadatos = (Select max(B1.FECHADATOS) FROM REPORTE_IDA B1))"
		else
		sql="select top " & cantidadxpagina & "  A.TERRITORIO, A.[Stock de cierre],A.[Ingresos del Mes],A.[Meta Mensual],A.[Meta Anual],A.[Avance Anual],A.[Avance Anual %],A.[Avance Mes],A.[Avance Mes %],A.REF,A.TRA,A.EJEC,A.EFE,A.OTRO,A.PART,A.PYME FROM Reporte_IDA A where A.FECHADATOS=(Select max(B.fechadatos) from Reporte_Ida B)"
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
										
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]='<%=RS.Fields("Territorio")%>';
				datos[tabla][<%=contador%>][1]='<%=rs.Fields("Stock de cierre")%>';
				datos[tabla][<%=contador%>][2]='<%=rs.Fields("Ingresos del mes")%>';
				datos[tabla][<%=contador%>][3]='<%=rs.Fields("Meta mensual")%>';
				datos[tabla][<%=contador%>][4]='<%=rs.Fields("Meta anual")%>';
				datos[tabla][<%=contador%>][5]='<%=rs.Fields("Avance Anual")%>';
				datos[tabla][<%=contador%>][6]='<%=rs.Fields("Avance Anual %")%>';
				datos[tabla][<%=contador%>][7]='<%=RS.Fields("Avance Mes")%>';
				datos[tabla][<%=contador%>][8]='<%=rs.Fields("Avance Mes %")%>';
				datos[tabla][<%=contador%>][9]='<%=rs.Fields("REF")%>';
				datos[tabla][<%=contador%>][10]='<%=rs.Fields("TRA")%>';
				datos[tabla][<%=contador%>][11]='<%=rs.Fields("EJEC")%>';
				datos[tabla][<%=contador%>][12]='<%=rs.Fields("EFE")%>';
				datos[tabla][<%=contador%>][13]='<%=rs.Fields("OTRO")%>';
				datos[tabla][<%=contador%>][14]='<%=rs.Fields("PART")%>';
				datos[tabla][<%=contador%>][15]='<%=rs.Fields("PYME")%>';

							
		<%
			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','','','','','','','','','','','','');


		    //Se escriben las opciones para los selects que contenga
		    posicionselect[tabla]=new Array();
		    nombreselect[tabla]=new Array();
		    opcionesvalor[tabla]=new Array();
		    opcionestexto[tabla]=new Array();
		    //Finaliza configuracion de tabla 0
		    
		    funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
		    funcionagrega[tabla]='agregar();';
		    

		</script> 
		<%if contador=0 then%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Oficina (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
			</tr>
			</table>
		<%else		
		%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Oficina (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar" align=middle></a>&nbsp;&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
			</tr>	
			</table>
			<div id="tabla0"> 
			</div>
		<%end if%>
		<input type="hidden" name="actualizarlista" value="">
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="pag" value="<%=pag%>">	
		</form>
		<script type="text/javascript">
			initTriStateCheckBox('tristateBox1', 'tristateBox1State', true);
		</script>
		<script language="javascript">
			inicio();
		</script>					
		</body>
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>		
		</html>		
		<%
		
				if expimp="1" then
					''Para Exportar a Excel
					''Paso Cero eliminar exportación anterior
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'Cód.Oficina','Descripción','Territorio','Grupo Ubigeo','Departamento','Provincia','Distrito','Activo'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select char(39) + A.codoficina,A.descripcion,B.descripcion as territorio,D.descripcion as Grupoubigeo, LTRIM(RTRIM(C.departamento)), C.provincia, C.distrito,CASE WHEN A.Activo=1 THEN 'Sí' ELSE 'No' END as FlagActivo " & _
								 "from CobranzaCM.dbo.Oficina A inner join CobranzaCM.dbo.Territorio B on A.codterritorio = B.codterritorio left outer join CobranzaCM.dbo.Ubigeo C on A.coddpto = C.coddpto and A.codprov = C.codprov and A.coddist = C.coddist left outer join CobranzaCM.dbo.GrupoUbigeo D on C.codgrupoubigeo = D.codgrupoubigeo " & filtrobuscador & " order by A.codoficina"
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



