f<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
''NOTA: Si se le da a un usuario el Perfil:Administrador por defecto se activa el flag: administrador y viceversa
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admagencia.asp") then
		buscador=obtener("buscador")
		buscaractivos=obtener("buscaractivos")
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
		var ventanaagencia;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanaagencia=global_popup_IWTSystem(ventanaagencia,"nuevoagencia.asp?vistapadre=" + window.name + "&paginapadre=admagencia.asp&codagencia=" + codigo,"Newagencia","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 350)/2 - 30) + ",height=380,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregar()
		{
			ventanaagencia=global_popup_IWTSystem(ventanaagencia,"nuevoagencia.asp?vistapadre=" + window.name + "&paginapadre=admagencia.asp","Newagencia","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 350)/2 - 30) + ",height=380,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=12;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('Codigo','Grupo','Nombre','Plaza','Grupo Efectividad','SubGrupo Efectividad','Razón Social','RUC','Teléfono','Dirección','Act','Editar');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(true, true,true , true , true , true, true,false, false, true, true , true);
		    anchocolumna[tabla] =  new Array( '1%',      '2%' , '3%' ,  '2%' , '2%'  , '2%',    '3%',      '2%' ,    '',    '2%',    '1%',    '1%');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left','left','center','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left','left','left','left','center','left','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1      ,-1   ,-1 ,-1     ,-1  ,-1 ,-1  ,-1 ,-1  ,-1 ,-1 );
		    formatofecha[tabla] =   new Array('' ,''  ,''  ,'' ,'','', '','', '','', '', '');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=codagencia-id- value=-c0->' + '<a href="javascript:modificar(-id-);">-valor-</a>';//+ objetodatos("text",tabla,"usuario","left","6","");
				objetofomulario[tabla][1]='<a href="javascript:modificar(-id-);">-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][4]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][6]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][7]='<a href=javascript:modificar("-id-");>-valor-</a>'; 
				objetofomulario[tabla][8]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][9]='<a href=javascript:modificar("-id-");>-valor-</a>'; 
				objetofomulario[tabla][10]=objetodatos("checkbox",tabla,"activo","","","");
				objetofomulario[tabla][11]='<a href=javascript:modificar("-id-");>Editar</a>';
	
										
					
		    filtrardatos[tabla]=0; //define si carga auto el filtro
		    filtrofomulario[tabla] = new Array();
		    tipofiltrofomulario[tabla] = new Array();
		    	filtrofomulario[tabla][0]='';
				filtrofomulario[tabla][1]='';
				filtrofomulario[tabla][2]='';
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
			filtrobuscador = " where (A.razonsocial like '%" & buscador & "%' or A.plaza like '%" & buscador & "%' or A.grupoefectividad like '%" & buscador & "%' or A.subgrupoefectividad like '%" & buscador & "%' or A.nombreSunat like '%" & buscador & "%' or A.rucSunat like '%" & buscador & "%' or A.telefono like '%" & buscador & "%' or A.direccion like '%" & buscador & "%' or B.descripcion like '%" & buscador & "%') "
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
		sql="select count(*) from agencia A left outer join grupoagencia B on A.codgrupoagencia = B.codgrupoagencia" & filtrobuscador 
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
		sql="select top " & cantidadxpagina & "A.codagencia, A.razonsocial, A.plaza,A.grupoefectividad,A.subgrupoefectividad,A.nombreSunat, A.rucSunat, A.telefono, A.direccion, A.activo, B.codgrupoagencia, B.descripcion as grupo " & _
			"from agencia A left outer join grupoagencia B on A.codgrupoagencia = B.codgrupoagencia where " & filtrobuscador1 & " A.codagencia not in (select top " & topnovisible & " A.codagencia from agencia A left outer join grupoagencia B on A.codgrupoagencia = B.codgrupoagencia " & filtrobuscador & " order by A.codagencia) order by A.codagencia" 
		else
		sql="select top " & cantidadxpagina & " A.codagencia, A.razonsocial,A.plaza,A.grupoefectividad,A.subgrupoefectividad,A.nombreSunat, A.rucSunat, A.telefono, A.direccion, A.activo, B.codgrupoagencia, B.descripcion as grupo " & _
			"from agencia A left outer join grupoagencia B on A.codgrupoagencia = B.codgrupoagencia " & filtrobuscador & " order by A.codagencia" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF
				if obtener("actualizarlista")<>"" and obtener("codagencia" & RS.Fields("codagencia"))<>"" then
										
					'''if obtener("razonsocial" & RS.Fields("codagencia"))<>"" then
					'''	razonsocial=obtener("razonsocial" & RS.Fields("codagencia"))
					'''else
					'''	razonsocial=""
					'''end if	
					
					'''if obtener("telefono" & RS.Fields("codagencia"))<>"" then
					'''	telefono=obtener("telefono" & RS.Fields("codagencia"))
					'''else
					'''	telefono=""
					'''end if	
					
					'''if obtener("direccion" & RS.Fields("codagencia"))<>"" then
					'''	direccion=obtener("direccion" & RS.Fields("codagencia"))
					'''else
					'''	direccion=""
					'''end if				
														
					if obtener("activo" & RS.Fields("codagencia"))<>"" then
						activo="1"
					else
						activo="0"
					end if						
										
						
						
						 	'''obtener("razonsocial" & RS.Fields("codagencia")) <> rs.Fields("razonsocial") or _
							'''obtener("telefono" & RS.Fields("codagencia")) <> rs.Fields("telefono") or _
							'''obtener("direccion" & RS.Fields("codagencia")) <> rs.Fields("direccion") or _
						if int(activo)<> rs.Fields("activo") then
							

							'''sql="update agencia set razonsocial='" & razonsocial & "',telefono='" & telefono & "',direccion='" & dieccion & "',activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codagencia=" & rs.Fields("codagencia") 
							sql="update agencia set activo=" & activo & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codagencia=" & rs.Fields("codagencia") 
							''response.write sql
							conn.Execute sql

						end if

				end if 

										
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]=<%=RS.Fields("codagencia")%>;
				datos[tabla][<%=contador%>][1]='<%=rs.Fields("grupo")%>';
			    datos[tabla][<%=contador%>][2]='<%=rs.Fields("razonsocial")%>';
			    datos[tabla][<%=contador%>][3]='<%=rs.Fields("plaza")%>';
				datos[tabla][<%=contador%>][4]='<%=rs.Fields("grupoefectividad")%>';
				datos[tabla][<%=contador%>][5]='<%=rs.Fields("subgrupoefectividad")%>';
			    datos[tabla][<%=contador%>][6]='<%=rs.Fields("nombreSunat")%>';
				datos[tabla][<%=contador%>][7]='<%=rs.Fields("rucSunat")%>';
				datos[tabla][<%=contador%>][8]='<%=rs.Fields("telefono")%>';
				datos[tabla][<%=contador%>][9]='<%=rs.Fields("direccion")%>';
				datos[tabla][<%=contador%>][10]=<%if obtener("actualizarlista")<>"" and obtener("codagencia" & RS.Fields("codagencia"))<>"" then%><%if int(activo)=1 then%>'checked'<%else%>' '<%end if%><%else%><%if rs.Fields("activo")=1 then%>'checked'<%else%>' '<%end if%><%end if%>;
				datos[tabla][<%=contador%>][11]='';
				
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
        objetosdebusqueda="<font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name='buscador' value='" & buscador & "' size=20 onkeypress='if(window.event.keyCode==13) buscar();'></font>&nbsp;<span id='tristateBox1' style='cursor: default;'>&nbsp;Activos<input type='hidden' id='tristateBox1State' name='buscaractivos' " & checkbuscactivos & "></span>"
		%>	
		
		<%if contador=0 then%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Agencia (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a></td>
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
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Agencia (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar" align=middle></a>&nbsp;&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
				<!--<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Agencias (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar" align=middle></a>&nbsp;&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a>&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a><%if expimp="1" then%>&nbsp;&nbsp;<a href='exportados/<%=nombrearchivo%>.xls','VerExport'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%>--></b></font></td>
				<!--<td bgcolor="#F5F5F5" align=middle width=200><font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size=20 onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
				<!--<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>-->
				<!--<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>-->
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
					consulta_exp="select 'Cod.Agencia','Grupo','Nombre','Plaza','Grupo Efectividad','SubGrupo Efectividad','Razón Social','RUC','Teléfono','Dirección','Activo'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select A.codagencia,B.descripcion,A.razonsocial,A.plaza,A.grupoefectividad,A.subgrupoefectividad,A.nombresunat,A.rucsunat,A.telefono, A.direccion,CASE WHEN A.Activo=1 THEN 'Sí' ELSE 'No' END as FlagActivo " & _
								 "from CobranzaCM.dbo.agencia A left outer join CobranzaCM.dbo.grupoagencia B on A.codgrupoagencia = B.codgrupoagencia  " & filtrobuscador & " order by A.codagencia" 
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



