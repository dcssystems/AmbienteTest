<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admoficina.asp") then
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
		  //Configuraci�n general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=9;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('Codigo','Descripci�n','Territorio','G.Ubigeo','Departamento','Provincia','Distrito','Act','Editar');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(true, true, true ,true,true,true,true,true,true);
		    anchocolumna[tabla] =  new Array( '5%', '15%', '15%' , '8%' ,'8%' ,'8%' ,'8%','3%' ,'3%');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','center','left');
		    aligndetalle[tabla] = new Array('center','left','left','left','left','left','left','center','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','center','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1 ,-1 ,-1 , -1 , -1 , '-1' , -1 , -1);
		    formatofecha[tabla] =   new Array(''  ,''   ,'' ,'','' ,'','','','');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=codoficina-id- value=-c0->' + '<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][1]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][4]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][6]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][7]=objetodatos("checkbox",tabla,"activo","","","");
				objetofomulario[tabla][8]='<a href=javascript:modificar("-id-");>Editar</a>';
					
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
							
		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		if buscador<>"" then
			filtrobuscador = " where (B.codterritorio + ' - ' + B.descripcion like '%" & buscador & "%' or A.descripcion like '%" & buscador & "%' or C.distrito like '%" & buscador & "%' or C.provincia like '%" & buscador & "%' or C.departamento like '%" & buscador & "%' or D.descripcion like '%" & buscador & "%') "
		end if
		
		
		select case buscaractivos
		case "0" : 
					checkbuscactivos="value='0'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.Activo=0"
		case "2" : 
					checkbuscactivos="value='2'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.Activo=1"
		case else: 
					checkbuscactivos="value='1'"
		end select	

		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if		
		
		contadortotal=0
		sql="select count(*) from Oficina A inner join Territorio B on A.codterritorio = B.codterritorio left outer join Ubigeo C on A.coddpto = C.coddpto and A.codprov = C.codprov and A.coddist = C.coddist left outer join grupoubigeo D on C.codgrupoubigeo = D.codgrupoubigeo " & filtrobuscador 
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
		sql="select top " & cantidadxpagina & "  A.codoficina, B.codterritorio,B.descripcion as territorio, A.descripcion, A.activo, A.coddpto, A.codprov, A.coddist, C.distrito, C.provincia, C.departamento, C.codgrupoubigeo, D.descripcion as Grupoubigeo from Oficina A inner join Territorio B on A.codterritorio = B.codterritorio left outer join Ubigeo C on A.coddpto = C.coddpto and A.codprov = C.codprov and A.coddist = C.coddist left outer join grupoubigeo D on C.codgrupoubigeo = D.codgrupoubigeo where " & filtrobuscador1 & " A.codoficina not in (select top " & topnovisible & " A.codoficina from Oficina A inner join Territorio B on A.codterritorio = B.codterritorio inner join Ubigeo C on A.coddpto = C.coddpto and A.codprov = C.codprov and A.coddist = C.coddist inner join grupoubigeo D on C.codgrupoubigeo = D.codgrupoubigeo " & filtrobuscador & " order by A.codoficina) order by A.codoficina" 
		else
		sql="select top " & cantidadxpagina & "  A.codoficina, B.codterritorio,B.descripcion as territorio, A.descripcion, A.activo, A.coddpto, A.codprov, A.coddist, C.distrito, C.provincia, C.departamento, C.codgrupoubigeo, D.descripcion as Grupoubigeo from Oficina A inner join Territorio B on A.codterritorio = B.codterritorio left outer join Ubigeo C on A.coddpto = C.coddpto and A.codprov = C.codprov and A.coddist = C.coddist left outer join grupoubigeo D on C.codgrupoubigeo = D.codgrupoubigeo " & filtrobuscador & " order by A.codoficina" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF 
				if obtener("actualizarlista")<>"" and obtener("codoficina" & RS.Fields("codoficina"))<>"" then
					
					if obtener("activo" & RS.Fields("codoficina"))<>"" then
						activo= "1"
					else
						activo= "0" 
					end if											
									
						if 	int(activo) <> rs.Fields("activo") then
							sql="update oficina set activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codoficina= '" & rs.Fields("codoficina") & "'" 
							''response.write sql
							conn.Execute sql
						end if
						
				end if 

										
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]='<%=RS.Fields("codoficina")%>';
				datos[tabla][<%=contador%>][1]='<%=rs.Fields("descripcion")%>';
				datos[tabla][<%=contador%>][2]='<%=rs.Fields("codterritorio") & " - " & rs.Fields("territorio")%>';
				datos[tabla][<%=contador%>][3]='<%=rs.Fields("Grupoubigeo")%>';
				datos[tabla][<%=contador%>][4]='<%=rs.Fields("departamento")%>';
				datos[tabla][<%=contador%>][5]='<%=rs.Fields("provincia")%>';
				datos[tabla][<%=contador%>][6]='<%=rs.Fields("distrito")%>';
				datos[tabla][<%=contador%>][7]=<%if obtener("actualizarlista")<>"" and obtener("codoficina" & RS.Fields("codoficina"))<>"" then%><%if int(activo)=1 then%>'checked'<%else%>' '<%end if%><%else%><%if rs.Fields("activo")=1 then%>'checked'<%else%>' '<%end if%><%end if%>;
				datos[tabla][<%=contador%>][8]='';							
		<%
			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','','','','','');


		    //Se escriben las opciones para los selects que contenga
		    posicionselect[tabla]=new Array();
		    nombreselect[tabla]=new Array();
		    opcionesvalor[tabla]=new Array();
		    opcionestexto[tabla]=new Array();
		    //Finaliza configuracion de tabla 0
		    
		    funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
		    funcionagrega[tabla]='agregar();';
		    
		<%
        objetosdebusqueda="<font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name='buscador' value='" & buscador & "' size=20 onkeypress='if(window.event.keyCode==13) buscar();'></font>&nbsp;<span id='tristateBox1' style='cursor: default;'>&nbsp;Activos<input type='hidden' id='tristateBox1State' name='buscaractivos' " & checkbuscactivos & "></span>"
		%>	

		</script> 
		<%if contador=0 then%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Oficina (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a></td>
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
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Oficina (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar" align=middle></a>&nbsp;&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>P�g.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
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
					''Paso Cero eliminar exportaci�n anterior
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'C�d.Oficina','Descripci�n','Territorio','Grupo Ubigeo','Departamento','Provincia','Distrito','Activo'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select char(39) + A.codoficina,A.descripcion,B.descripcion as territorio,D.descripcion as Grupoubigeo, LTRIM(RTRIM(C.departamento)), C.provincia, C.distrito,CASE WHEN A.Activo=1 THEN 'S�' ELSE 'No' END as FlagActivo " & _
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
		alert("Ud. No tiene autorizaci�n para este proceso.");
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



