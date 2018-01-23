<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
''Response.ContentType="application/download"
''Response.Redirect "importados/respuestas63.txt"
if session("codusuario")<>"" then
	conectar
	if permisofacultad("logcarga.asp") then
		buscador=obtener("buscador")
		buscaractivos=obtener("buscaractivos")
	''Codigo exp excel - se repite
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
		
		sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar' or descripcion='RutaWebExportar'"
		consultar sql,RS
		RS.Filter=" descripcion='RutaFisicaExportar'"
		RutaFisicaExportar=RS.Fields(1)
		RS.Filter=" descripcion='RutaWebExportar'"
		RutaWebExportar=RS.Fields(1)				
		RS.Filter=""					
		RS.Close

        sql="select descripcion,valortexto1 from parametro where descripcion='RutaWebUpload' or descripcion='RutaFisicaUpload'"
        consultar sql,RS
		RS.Filter=" descripcion='RutaWebUpload'"
		RutaWebUpload=RS.Fields(1)
		RS.Filter=" descripcion='RutaFisicaUpload'"
		RutaFisicaUpload=RS.Fields(1)				
		RS.Filter=""					
        RS.Close	
			
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<script language=javascript src="scripts/TablaDinamica.js"></script>
		<script language=javascript>
		var ventanalogcarga;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanalogcarga=global_popup_IWTSystem(ventanalogcarga,"nuevologcarga.asp?vistapadre=" + window.name + "&paginapadre=admlogcarga.asp&codlogcarga=" + codigo,"Newlogcarga","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 200)/2 - 30) + ",height=200,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}	
			
		function agregar()
		{
			ventanalogcarga=global_popup_IWTSystem(ventanalogcarga,"nuevologcarga.asp?vistapadre=" + window.name + "&paginapadre=admlogcarga.asp","Newlogcarga","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 200)/2 - 30) + ",height=200,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}
		function abrirarchivo(nomarch) 
		{
			window.open("descargararchivo.asp?nomarch=" + nomarch,"_self");
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
		    ascendente[tabla]=false;
		    nrocolumnas[tabla]=12;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('Codigo','Proceso','Ruta','Estado','Usuario','Agencia','Registro','Observaciones','Inicio','Fin','Desc','Error');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(true, true, true ,true,true , true, true ,true,true,true,true,true);
		    anchocolumna[tabla] =  new Array( '2%', '3%', '15%' , '7%' ,'3%', '10%', '8%' , '' ,'4%','4%','4%','4%');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left','left','left','left','left','center','center');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1 ,-1 ,-1 , -1 ,-1 ,-1 ,-1 , -1 , -1 , -1 ,-1 );
		    formatofecha[tabla] =   new Array(''  ,''   ,'' ,'','',''   ,'dd/mm/aaaa&nbsp;HH:MI:SS' ,'','dd/mm&nbsp;HH:MI:SS' ,'dd/mm&nbsp;HH:MI:SS','' ,'');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='-valor-';
				objetofomulario[tabla][1]='-valor-';
				objetofomulario[tabla][2]='-valor-';
				objetofomulario[tabla][3]='-valor-';
				objetofomulario[tabla][4]='-valor-';
				objetofomulario[tabla][5]='-valor-';
				objetofomulario[tabla][6]='-valor-';
				objetofomulario[tabla][7]='-valor-';
				objetofomulario[tabla][8]='-valor-';
				objetofomulario[tabla][9]='-valor-';
				objetofomulario[tabla][10]='-valor-';
				objetofomulario[tabla][11]='-valor-';
	
				// target="DescArch"   onclick="window.open(this.href,3&39window3&39,3&39params3&39);return false" 3&39=' onclick="abrir(this.href,3&39archivo.txt3&39);return false"
										
					
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
				filtrofomulario[tabla][9]='';
				filtrofomulario[tabla][10]='';
				filtrofomulario[tabla][11]='';
		
				
										
					
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
			
		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		if buscador<>"" then
			filtrobuscador = " where (d.razonsocial like '%" & buscador & "%' or a.codcarga like '%" & buscador & "%' or a.proceso like '%" & buscador & "%' or a.rutaarchivo like '%" & buscador & "%' or b.descripcion like '%" & buscador & "%' or c.usuario like '%" & buscador & "%' or a.fecharegistra like '%" & buscador & "%' or a.observaciones like '%" & buscador & "%' or a.fechainicio like '%" & buscador & "%' or a.fechafin like '%" & buscador & "%') "
		end if
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if
				
		contadortotal=0
		sql="select count(*) from cargaarchivo a inner join estadocarga b on a.codestadocarga = b.codestadocarga inner join usuario c on a.usuarioregistra = c.codusuario left outer join agencia d on c.codagencia = d.codagencia " & filtrobuscador 
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
		sql="select top " & cantidadxpagina & " d.razonsocial, a.codcarga, a.proceso, a.rutaarchivo, convert(varchar,a.fecharegistra,112) as fregistra,  CASE WHEN b.codestadocarga=3 OR b.codestadocarga=4 then NULL else b.codestadocarga end as codestadocarga, b.descripcion, c.usuario, a.fecharegistra, a.observaciones, a.fechainicio, a.fechafin from cargaarchivo a inner join estadocarga b on a.codestadocarga = b.codestadocarga inner join usuario c on a.usuarioregistra = c.codusuario left outer join agencia d on c.codagencia = d.codagencia where " & filtrobuscador1 & " a.codcarga not in (select top " & topnovisible & " a.codcarga from cargaarchivo a inner join estadocarga b on a.codestadocarga = b.codestadocarga inner join usuario c on a.usuarioregistra = c.codusuario " & filtrobuscador & " order by a.codcarga desc) order by a.codcarga desc" 
		else
		sql="select top " & cantidadxpagina & " d.razonsocial, a.codcarga, a.proceso, a.rutaarchivo, convert(varchar,a.fecharegistra,112) as fregistra,  CASE WHEN b.codestadocarga=3 OR b.codestadocarga=4 then NULL else b.codestadocarga end as codestadocarga, b.descripcion, c.usuario, a.fecharegistra, a.observaciones, a.fechainicio, a.fechafin from cargaarchivo a inner join estadocarga b on a.codestadocarga = b.codestadocarga inner join usuario c on a.usuarioregistra = c.codusuario left outer join agencia d on c.codagencia = d.codagencia " & filtrobuscador & " order by a.codcarga desc" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF 
				''rutaarchivo=RS.Fields("rutaarchivo")
				''nombrearchivo=""
				''if len(rutaarchivo)>0 then
				''	for i=1 to len(rutaarchivo)
				''		if mid(rutaarchivo,len(rutaarchivo)- i + 1,1)<>"\" then
				''			nombrearchivo=mid(rutaarchivo,len(rutaarchivo) - i + 1,1) & nombrearchivo
					''	else 
					''		exit for
					''	end if
				''	next
				''end if
		%>
				datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]=<%=RS.Fields("codcarga")%>;
				datos[tabla][<%=contador%>][1]='<%=RS.Fields("proceso")%>';
				datos[tabla][<%=contador%>][2]='<%=replace(RS.Fields("rutaarchivo"),"\","\\")%>';
				datos[tabla][<%=contador%>][3]='<%=iif(isNull(rs.Fields("descripcion")),"",rs.Fields("descripcion"))%>';
				datos[tabla][<%=contador%>][4]='<%=iif(isNull(rs.Fields("usuario")),"",rs.Fields("usuario"))%>';
				datos[tabla][<%=contador%>][5]='<%=iif(isNull(rs.Fields("razonsocial")),"",rs.Fields("razonsocial"))%>';
				datos[tabla][<%=contador%>][6]=<%if not IsNull(RS.Fields("fecharegistra")) then%>new Date(<%=Year(RS.Fields("fecharegistra"))%>,<%=Month(RS.Fields("fecharegistra"))-1%>,<%=Day(RS.Fields("fecharegistra"))%>,<%=Hour(RS.Fields("fecharegistra"))%>,<%=Minute(RS.Fields("fecharegistra"))%>,<%=Second(RS.Fields("fecharegistra"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][7]='<%=RS.Fields("observaciones")%>';
				datos[tabla][<%=contador%>][8]=<%if not IsNull(RS.Fields("fechainicio")) then%>new Date(<%=Year(RS.Fields("fechainicio"))%>,<%=Month(RS.Fields("fechainicio"))-1%>,<%=Day(RS.Fields("fechainicio"))%>,<%=Hour(RS.Fields("fechainicio"))%>,<%=Minute(RS.Fields("fechainicio"))%>,<%=Second(RS.Fields("fechainicio"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][9]=<%if not IsNull(RS.Fields("fechafin")) then%>new Date(<%=Year(RS.Fields("fechafin"))%>,<%=Month(RS.Fields("fechafin"))-1%>,<%=Day(RS.Fields("fechafin"))%>,<%=Hour(RS.Fields("fechafin"))%>,<%=Minute(RS.Fields("fechafin"))%>,<%=Second(RS.Fields("fechafin"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][10]='<a href="<%=RutaWebUpload & "/" & replace(RS.Fields("rutaarchivo"),RutaFisicaUpload & "\","")%>" target="T_New"><img src="imagenes/descargarpeq.png" border=0 alt="Descargar Archivo" title="Descargar Archivo"></a>';
				datos[tabla][<%=contador%>][11]='<%if not IsNull(RS.Fields("codestadocarga")) then%><a href="<%=RutaWebExportar%>/<%=RS.Fields("fregistra")%>_Error_<%=RS.Fields("codcarga")%>.rar" target="T_New"><img src="imagenes/descargarpeq.png" border=0 alt="Descargar Error" title="Descargar Error"></a><%end if%>';
		<%
			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','','','','','','','',''); 


		    //Se escriben las opciones para los selects que contenga
		    posicionselect[tabla]=new Array();
		    nombreselect[tabla]=new Array();
		    opcionesvalor[tabla]=new Array();
		    opcionestexto[tabla]=new Array();
		    //Finaliza configuracion de tabla 0
		    
			    
		    funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
		    funcionagrega[tabla]='agregar();';

		</script> 
		
		
		
		
		<%
        objetosdebusqueda="<font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name='buscador' value='" & buscador & "' size=20 onkeypress='if(window.event.keyCode==13) buscar();'></font></span>"
		%>	
				
		<%if contador=0 then%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Log Cargas (0) - No hay registros.</b></font></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
			</tr>
			</table>
		<%else		
		%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Log Cargas (<%=contadortotal%>)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
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
		<script language="javascript">
			inicio();
		</script>	
		</body>
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>		
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
				''Codigo','Proceso','Ruta','Estado','Usuario','Fecha registro','Observaciones','Fecha inicio','Fecha fin');
					''Para Exportar a Excel
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'Cod.Carga','Proceso','Ruta','Estado','Usuario','Fecha registro','Observaciones','Fecha inicio','Fecha fin' "
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select a.codcarga, a.proceso, a.rutaarchivo, b.descripcion, c.usuario, convert(varchar,a.fecharegistra,103) + ' ' + convert(varchar,a.fecharegistra,108) AS fecharegistra, a.observaciones, convert(varchar,a.fechainicio,103) + ' ' + convert(varchar,a.fechainicio,108) AS fechainicio, convert(varchar,a.fechafin,103) + ' ' + convert(varchar,a.fechafin,108) AS fechafin " & _
								 "from CobranzaCM.dbo.cargaarchivo a inner join CobranzaCM.dbo.estadocarga b on a.codestadocarga = b.codestadocarga inner join CobranzaCM.dbo.usuario c on a.usuarioregistra = c.codusuario " & filtrobuscador & " order by A.codcarga desc" 
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



