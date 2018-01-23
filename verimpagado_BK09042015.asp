<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admimpagado.asp") then
		codigocentral=obtener("codigocentral")
		contrato=obtener("contrato")
		fechadatos=obtener("fechadatos")
		buscador=obtener("buscador")
		''fechadatos=mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2)
		fechagestion=obtener("fechagestion")
		fechagestionini=obtener("fechagestionini")
		fechagestionfin=obtener("fechagestionfin")
		''fechagestion=mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)
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
		
		sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload' or descripcion='RutaWebUpload'"
		consultar sql,RS3
		RS3.Filter=" descripcion='RutaFisicaUpload'"
		RutaFisicaUpload=RS3.Fields(1)
		RS3.Filter=" descripcion='RutaWebUpload'"
		RutaWebUpload=RS3.Fields(1)				
		RS3.Filter=""
		RS3.Close	
		
		sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
		consultar sql,RS	
		maxfechagestion=rs.fields(0)
		RS.Close	
		
		if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
			vistabusqueda="VERULTIMPAGADO"
		else
			vistabusqueda="VERIMPAGADO"
		end if				
		
		if obtener("direccioneliminar")<>"" then
			sql="Update direccionnueva set activo=0,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codigocentral='" & codigocentral & "' and coddireccionnueva=" & obtener("direccioneliminar")
			conn.execute sql
		end if
		
		if obtener("telefonoeliminar")<>"" then
			sql="Update telefononuevo set activo=0,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codigocentral='" & codigocentral & "' and codtelefononuevo=" & obtener("telefonoeliminar")
			conn.execute sql
		end if
		
		if obtener("emaileliminar")<>"" then
			sql="Update emailnuevo set activo=0,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codigocentral='" & codigocentral & "' and codemailnuevo=" & obtener("emaileliminar")
			conn.execute sql
		end if
		

		
		'''sql="select diasvencimiento,nombre,agencia,tipodocumento,numdocumento,direccion,departamento,provincia,distrito,tipofono1,prefijo1,fono1,extension1,tipofono2,prefijo2,fono2,extension2,tipofono3,prefijo3,fono3,extension3,tipofono4,prefijo4,fono4,extension4,tipofono5,prefijo5,fono5,extension5,email,(select top 1 Y.descripcion from MarcaCliente X inner join Marca Y on X.codigocentral=A.codigocentral and X.codmarca=Y.codmarca and X.activo=1 order by X.codmarcacliente desc) as Marca from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & fechadatos & "' order by DiasVencimiento desc"
		''sql="select *,(select top 1 Y.descripcion from MarcaCliente X inner join Marca Y on X.codigocentral=A.codigocentral and X.codmarca=Y.codmarca and X.activo=1 order by X.codmarcacliente desc) as Marca,(select count(distinct fechavencimiento) from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos) as NumCuotas,(select top 1 divisa from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos and divisa<>A.divisa) as DivisaDif from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & fechadatos & "' order by DiasVencimiento desc,saldohoy desc"
		''sql="select *,(select top 1 Y.descripcion from MarcaCliente X inner join Marca Y on X.codigocentral=A.codigocentral and X.codmarca=Y.codmarca and X.activo=1 order by X.codmarcacliente desc) as Marca from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' order by DiasVencimiento desc,saldohoy desc"
		sql="select * from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' order by DiasVencimiento desc,saldohoy desc"
		''Response.Write sql
		consultar sql,RS1	
		if not RS1.EOF then
			nombre=RS1.Fields("nombre")
			agencia=RS1.Fields("agencia")
			tipodocumento=RS1.Fields("tipodocumento")
			numdocumento=RS1.Fields("numdocumento")
			nrocontratos=RS1.RecordCount
			codmarca=RS1.Fields("codmarca")
			marca=RS1.Fields("marca")
			MaxDias=RS1.Fields("DiasVencimiento")
			direccion=RS1.Fields("direccion")
			departamento=RS1.Fields("departamento")
			provincia=RS1.Fields("provincia")
			distrito=RS1.Fields("distrito")
			tipofono1=RS1.Fields("tipofono1")
			prefijo1=RS1.Fields("prefijo1")
			fono1=RS1.Fields("fono1")
			extension1=RS1.Fields("extension1")
			tipofono2=RS1.Fields("tipofono2")
			prefijo2=RS1.Fields("prefijo2")
			fono2=RS1.Fields("fono2")
			extension2=RS1.Fields("extension2")
			tipofono3=RS1.Fields("tipofono3")
			prefijo3=RS1.Fields("prefijo3")
			fono3=RS1.Fields("fono3")
			extension3=RS1.Fields("extension3")
			tipofono4=RS1.Fields("tipofono4")
			prefijo4=RS1.Fields("prefijo4")
			fono4=RS1.Fields("fono4")
			extension4=RS1.Fields("extension4")
			tipofono5=RS1.Fields("tipofono5")
			prefijo5=RS1.Fields("prefijo5")
			fono5=RS1.Fields("fono5")
			extension5=RS1.Fields("extension5")
			email=RS1.Fields("email")
		end if
		
		
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<title>Ver Impagado</title>
		<script language=javascript src="scripts/TablaDinamica.js"></script>
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
		var nuevoemail;
		var nuevogestion;
		var ventanagestion;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(c0)
		{
			ventanagestion=global_popup_IWTSystem(ventanagestion,"verseguimiento.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado3.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codrespgestion=" + c0,"NewGestion","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=300,width=" + (screen.width/2 + 100) + ",left=" + (screen.width/4 - 50) + ",resizable=yes");
		}		
		function agregardir()
		{
			nuevodir=global_popup_IWTSystem(nuevodir,"adicionardireccion.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado3.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newdir","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 + 250) + ",left=" + (screen.width/4) + ",resizable=yes");
		}
		function eliminardir(elimdir)
		{
			if(confirm("¿Está Seguro de Eliminar la Dirección del Cliente?"))
			{						
				document.formula.direccioneliminar.value=elimdir;
				document.formula.submit();
			}			
		}		
		function agregartelf()
		{
			nuevotelf=global_popup_IWTSystem(nuevotelf,"adicionartelefono.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado2.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newtelf","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
		}
		function eliminartelf(elimtelf)
		{
			if(confirm("¿Está Seguro de Eliminar el Teléfono del Cliente?"))
			{						
				document.formula.telefonoeliminar.value=elimtelf;
				document.formula.submit();
			}			
		}				
		function agregaremail()
		{
			nuevoemail=global_popup_IWTSystem(nuevoemail,"adicionaremail.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado2.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newemail","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
		}		
		function eliminaremail(elimmail)
		{
			if(confirm("¿Está Seguro de Eliminar el E-Mail del Cliente?"))
			{						
				document.formula.emaileliminar.value=elimmail;
				document.formula.submit();
			}			
		}		
		function agregargestion()
		{
			nuevogestion=global_popup_IWTSystem(nuevogestion,"adicionarseguimiento.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado2.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","NewGestion","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=300,width=" + (screen.width/2 + 100) + ",left=" + (screen.width/4 - 50) + ",resizable=yes");
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
		
		<script language=javascript>
			rutaimgcab="imagenes/"; 
		  //Configuración general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=1;
		    ascendente[tabla]=false;
		    nrocolumnas[tabla]=9;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '2';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('CodRespGestion','Fecha','Contrato','Agencia','Gestion','Comentario','F.Promesa','Direccion/Telefono','Adj');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array( false, true, true, true, true,true,true,true,true);
		    anchocolumna[tabla] =  new Array( '' , '5%' ,  '1%',  '3%',  '6%',  '8%','3%','8%','1%');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left','left','left','center');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1,-1,-1 ,-1 ,-1,-1,-1,-1);
		    formatofecha[tabla] =   new Array('','dd/mm/aaaa HH:MI'  ,'' ,''  ,'' ,'','dd/mm/aaaa','','');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=CodRespGestion-id- value=-c0-><a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][1]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][4]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][6]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][7]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][8]='-valor-';
				
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
		
		filtrobuscador = " where A.codigocentral='" & codigocentral & "'"
				
		if fechagestionini<>"" or fechagestionfin<>"" then
		    if fechagestionini<>"" then
                filtrobuscador = filtrobuscador & " and A.fhgestionado>='" & mid(fechagestionini,7,4) & mid(fechagestionini,4,2) & mid(fechagestionini,1,2) & "'"
            end if
            if fechagestionfin<>"" then

                 filtrobuscador = filtrobuscador & " and A.fhgestionado<='" & mid(fechagestionfin,7,4) & mid(fechagestionfin,4,2) & mid(fechagestionfin,1,2) & "'"
            end if
		end if
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if	
		
		contadortotal=0
		sql="select count(*) from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia " & filtrobuscador
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
		    sql="select top " & cantidadxpagina & " A.codrespgestion,A.fhgestionado,A.contrato,A.comentario,A.fechapromesa,A.fono,A.tipofono,A.prefijo,A.extension,A.direccion,A.ficherogestion,C.RazonSocial,B.Descripcion,UB.distrito,UB.provincia,UB.departamento from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia LEFT OUTER JOIN Ubigeo AS UB ON UB.coddpto = A.coddpto AND UB.codprov = A.codprov AND UB.coddist = A.coddist where " & filtrobuscador1 & " A.fhgestionado not in (select top " & topnovisible & " A.fhgestionado from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia " & filtrobuscador & " order by A.fhgestionado desc) order by A.fhgestionado desc" 
		else
		    sql="select top " & cantidadxpagina & " A.codrespgestion,A.fhgestionado,A.contrato,A.comentario,A.fechapromesa,A.fono,A.tipofono,A.prefijo,A.extension,A.direccion,A.ficherogestion,C.RazonSocial,B.Descripcion,UB.distrito,UB.provincia,UB.departamento from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia LEFT OUTER JOIN Ubigeo AS UB ON UB.coddpto = A.coddpto AND UB.codprov = A.codprov AND UB.coddist = A.coddist " & filtrobuscador & " order by A.fhgestionado desc" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		    Do while not RS.EOF
		%>
			datos[tabla][<%=contador%>] = new Array();
			    datos[tabla][<%=contador%>][0]='<%=RS.Fields("codrespgestion")%>';
				datos[tabla][<%=contador%>][1]=<%if not IsNull(RS.Fields("fhgestionado")) then%>new Date(<%=Year(RS.Fields("fhgestionado"))%>,<%=Month(RS.Fields("fhgestionado"))-1%>,<%=Day(RS.Fields("fhgestionado"))%>,<%=Hour(RS.Fields("fhgestionado"))%>,<%=Minute(RS.Fields("fhgestionado"))%>,<%=Second(RS.Fields("fhgestionado"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][2]='<%=rs.Fields("Contrato")%>';
			    datos[tabla][<%=contador%>][3]='<%=rs.Fields("RazonSocial")%>';
				datos[tabla][<%=contador%>][4]='<%=RS.Fields("descripcion")%>';
				datos[tabla][<%=contador%>][5]='<%if len(trim(RS.Fields("comentario")))<=25 then%><%=trim(RS.Fields("comentario"))%><%else%><%=mid(trim(RS.Fields("comentario")),1,25) & "..."%><%end if%>';
				datos[tabla][<%=contador%>][6]=<%if not IsNull(RS.Fields("fechapromesa")) then%>new Date(<%=Year(RS.Fields("fechapromesa"))%>,<%=Month(RS.Fields("fechapromesa"))-1%>,<%=Day(RS.Fields("fechapromesa"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][7]='<%if trim(RS.Fields("fono"))<>"" then%><%=RS.Fields("tipofono") & " - " & RS.Fields("prefijo") & " - " & RS.Fields("fono")%><%if len(trim(RS.Fields("extension"))) and RS.Fields("extension")<>"0000" then%><%=" - " & RS.Fields("extension")%><%end if%><%else%><%if trim(RS.Fields("Direccion"))<>"" then%><%=RS.Fields("Direccion") & " - " & RS.Fields("Distrito") & " - " & RS.Fields("Provincia") & " - " & RS.Fields("Departamento")%><%end if%><%end if%>';
				datos[tabla][<%=contador%>][8]='<%if not IsNull(RS.Fields("ficherogestion")) then%><a href="<%=RutaWebUpload%>/<%=RS.Fields("ficherogestion")%>" target="T_New"><img src="imagenes/descargarpeq.png" border=0 alt="Descargar Archivo" title="Descargar Archivo"></a><%end if%>';
				
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
		
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td align="center" height="22"><font size=2 face=Arial color="#FFFFFF"><b>Módulo Gestión On-Line Impagado</b></font></td>
			</tr>
			</table>
			<table width=100% cellpadding=2 cellspacing=2 border=0>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Fecha de Asignación:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=fechagestion%></b></font></td>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Agencia:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=agencia%></b></font></td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Código Central:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codigocentral%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Nombres:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=nombre%></b></font></td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;N° Documento:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=tipodocumento%> - <%=numdocumento%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;N° Contratos:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=nrocontratos%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Marca:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%if codmarca<>"" then%><%=codmarca & " - " & marca%><%end if%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Días Vencimiento:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=MaxDias%></b></font></td>
			</tr>											
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Direcciones:</font></td>
						<td align=right><a href="javascript:agregardir();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar Dirección" title="Agregar Dirección"></a>&nbsp;</td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
						<script language=javascript>
						function visualizardir()
						{
							var filas = document.getElementById('tabladirecciones').rows.length;
							//if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
							if (document.getElementById('imagendir').title=="Mostrar")
							{
								document.getElementById('imagendir').title="Ocultar";
								document.getElementById('imagendir').alt="Ocultar";
								document.getElementById('imagendir').src="imagenes/ocultar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tabladirecciones').rows[i].style.display = '';
								}
							}
							else
							{
								document.getElementById('imagendir').title="Mostrar";
								document.getElementById('imagendir').alt="Mostrar";
								document.getElementById('imagendir').src="imagenes/mostrar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tabladirecciones').rows[i].style.display = 'none';
								}							
							}
						}
						</script>
				 		<table id="tabladirecciones" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizardir();"><img id="imagendir" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Dirección</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Distrito</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Provincia</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Departamento</b></font></td>
						</tr>					 		
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=direccion%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=distrito%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=provincia%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=departamento%></font></td>
						</tr>	
						<%
						sql="select A.coddireccionnueva,A.direccion,B.departamento,B.provincia,B.distrito from DireccionNueva A left outer join Ubigeo B on A.coddpto=B.coddpto and A.codprov=B.codprov and A.coddist=B.coddist where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.coddireccionnueva desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminardir('<%=RS.Fields("coddireccionnueva")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar Dirección" title="Eliminar Dirección"></a></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("direccion")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("distrito")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("provincia")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("departamento")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>						
						</table>
						<%if obtener("agreguedir")<>"" then%>
						<script language=javascript>
							visualizardir();
						</script>
						<%end if%>
				 </td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Teléfonos:</font></td>
						<td align=right><a href="javascript:agregartelf();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar Teléfono" title="Agregar Teléfono"></a>&nbsp;</td>
					</tr>
					</table>						
				 </td>
				 <td colspan=3>
						<script language=javascript>
						function visualizartelf()
						{
							var filas = document.getElementById('tablatelefonos').rows.length;
							if (document.getElementById('imagentel').title=="Mostrar")
							{
								document.getElementById('imagentel').title="Ocultar";
								document.getElementById('imagentel').alt="Ocultar";
								document.getElementById('imagentel').src="imagenes/ocultar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablatelefonos').rows[i].style.display = '';
								}
							}
							else
							{
								document.getElementById('imagentel').title="Mostrar";
								document.getElementById('imagentel').alt="Mostrar";
								document.getElementById('imagentel').src="imagenes/mostrar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablatelefonos').rows[i].style.display = 'none';
								}							
							}
						}
						</script>
				 		<table id="tablatelefonos" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizartelf();"><img id="imagentel" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td width=38 bgcolor="#BEE8FB" align="center"><font size=2 face=Arial color=#00529B><b>&nbsp;Tipo</b></font></td>
							<td width=40 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Pref</b></font></td>
							<td width=80 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Teléfono</b></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Extensión</b></font></td>
						</tr>					 		
						<%if IsNumeric(fono1) then%>
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono1%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension1<>"0000" then%><%=extension1%><%end if%></font></td>
						</tr>	
						<%end if%>
						<%if IsNumeric(fono2) then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono2%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension2<>"0000" then%><%=extension2%><%end if%></font></td>
						</tr>	
						<%end if%>						
						<%if IsNumeric(fono3) then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono3%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension3<>"0000" then%><%=extension3%><%end if%></font></td>
						</tr>	
						<%end if%>
						<%if IsNumeric(fono4) then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono4%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension4<>"0000" then%><%=extension4%><%end if%></font></td>
						</tr>	
						<%end if%>
						<%if IsNumeric(fono5) then%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=tipofono5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=prefijo5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=fono5%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%if extension5<>"0000" then%><%=extension5%><%end if%></font></td>
						</tr>	
						<%end if%>																		
						<%		
						sql="select codtelefononuevo,codtipotelefono,prefijo,fono,extension from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminartelf('<%=RS.Fields("codtelefononuevo")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar Teléfono" title="Eliminar Teléfono"></a></font></td>
							<td bgcolor="#E9F8FE" align="center"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("codtipotelefono")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("prefijo")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("fono")%></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("extension")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>							
						</table>
						<%if obtener("agreguetelf")<>"" then%>
						<script language=javascript>
							visualizartelf();
						</script>
						<%end if%>
				 </td>
			</tr>							
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;E-mail:</font></td>
						<td align=right><a href="javascript:agregaremail();"><img src="imagenes/agregar.png" height=16 border=0 style="align: right; vertical-align: bottom;" alt="Agregar E-mail" title="Agregar E-mail"></a>&nbsp;</td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
						<script language=javascript>
						function visualizaremail()
						{
							var filas = document.getElementById('tablaemails').rows.length;
							//if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
							if (document.getElementById('imagenemail').title=="Mostrar")
							{
								document.getElementById('imagenemail').title="Ocultar";
								document.getElementById('imagenemail').alt="Ocultar";
								document.getElementById('imagenemail').src="imagenes/ocultar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablaemails').rows[i].style.display = '';
								}
							}
							else
							{
								document.getElementById('imagenemail').title="Mostrar";
								document.getElementById('imagenemail').alt="Mostrar";
								document.getElementById('imagenemail').src="imagenes/mostrar.png";
								for (i = 2; i < filas; i++)
								{
									document.getElementById('tablaemails').rows[i].style.display = 'none';
								}							
							}
						}
						</script>
				 		<table id="tablaemails" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizaremail();"><img id="imagenemail" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;E-mail</b></font></td>
						</tr>					 		
						<tr>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=email%></font></td>
						</tr>	
						<%		
						sql="select A.codemailnuevo,A.email from EmailNuevo A where A.codigocentral='" & codigocentral & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codemailnuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:eliminaremail('<%=RS.Fields("codemailnuevo")%>');"><img src="imagenes/eliminar.png" border=0 alt="Eliminar E-Mail" title="Eliminar E-Mail"></a></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("email")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>						
						</table>
						<%if obtener("agregueemail")<>"" then%>
						<script language=javascript>
							visualizaremail();
						</script>
						<%end if%>
				 </td>
			</tr>				
			</table>
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;Obligaciones del Cliente</b></font></td>
			</tr>
			</table>		
			<script language=javascript>
			function visualizarcuotas(numcont)
			{
				var filas = document.getElementById('tablacontratos').rows.length;
				//if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
				if (document.getElementById('imagencuota' + numcont).title=="Mostrar")
				{
					document.getElementById('imagencuota' + numcont).title="Ocultar";
					document.getElementById('imagencuota' + numcont).alt="Ocultar";
					document.getElementById('imagencuota' + numcont).src="imagenes/ocultar.png";
					for (i = 1; i < filas; i++)
					{
						if(document.getElementById('tablacontratos').rows[i].id==numcont) document.getElementById('tablacontratos').rows[i].style.display = '';
					}
				}
				else
				{
					document.getElementById('imagencuota' + numcont).title="Mostrar";
					document.getElementById('imagencuota' + numcont).alt="Mostrar";
					document.getElementById('imagencuota' + numcont).src="imagenes/mostrar.png";
					for (i = 1; i < filas; i++)
					{
						if(document.getElementById('tablacontratos').rows[i].id==numcont) document.getElementById('tablacontratos').rows[i].style.display = 'none';
					}							
				}
			}
			</script>				
			<table width=100% id="tablacontratos" cellpadding=1 cellspacing=1 border=0>
			<tr bgcolor="#007DC5">
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Cuotas</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>N° Contrato</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Producto</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>F.Incump</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>DA</b></font></td>				  
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Mon</b></font></td>
					<td align="center"><font size=2 face=Arial color="#FFFFFF"><b>Total</b></font></td>				  
			</tr>
			<%
			Do While not RS1.EOF
				sql="select * from CuotaDiario where contrato='" & RS1.Fields("Contrato") & "' and fechadatos='" & year(RS1.Fields("FechaDatos")) & right("00" & month(RS1.Fields("FechaDatos")),2) & right("00" & day(RS1.Fields("FechaDatos")),2) & "' order by fechavencimiento,divisa"
				consultar sql,RS2
				nrocuotas=0
				DivisaDif=""
				MontoTotalDivisa=0
				MontoTotalDivisaDif=0
				fechavencimiento=""
				Do While Not RS2.EOF
					if fechavencimiento<>RS2.Fields("FechaVencimiento") then
						fechavencimiento=RS2.Fields("FechaVencimiento")
						nrocuotas=nrocuotas + 1
					end if
					
					if RS2.Fields("Divisa")<>RS1.Fields("Divisa") then
						DivisaDif=RS2.Fields("Divisa")
						if RS1.Fields("codproducto")="50" then
							MontoTotalDivisaDif=MontoTotalDivisaDif + RS2.Fields("Capital") + RS2.Fields("Interes") + RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("Seguro") + RS2.Fields("InteresMoratorio") + RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
						else
							MontoTotalDivisaDif=MontoTotalDivisaDif + RS2.Fields("Capital") + RS2.Fields("Interes") + RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("InteresMoratorio") + RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
						end if	
					else
						if RS1.Fields("codproducto")="50" then
							MontoTotalDivisa=MontoTotalDivisa + RS2.Fields("Capital") + RS2.Fields("Interes") + RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("Seguro") + RS2.Fields("InteresMoratorio") + RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
						else
							MontoTotalDivisa=MontoTotalDivisa + RS2.Fields("Capital") + RS2.Fields("Interes") + RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("InteresMoratorio") + RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
						end if	
					end if					
				RS2.MoveNext
				Loop

				'',(select count(distinct fechavencimiento) from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos) as NumCuotas,(select top 1 divisa from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos and divisa<>A.divisa) as DivisaDif 
			%>
			<tr bgcolor="#E9F8FE">
					<td valign="top" align="center"><table cellspacing=0 cellpadding=0 border=0><tr><td><a href="javascript:visualizarcuotas('<%=RS1.Fields("Contrato")%>');"><img id="imagencuota<%=RS1.Fields("contrato")%>" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></td><td><font size=2 face=Arial color=#00529B>&nbsp;<%=nrocuotas%></font></td></tr></table></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Contrato")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Producto")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("FechaIncumplimiento")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("DiasVencimiento")%></font></td>				  
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=RS1.Fields("divisa")%></font></td>
					<td valign="top" align="center"><font size=2 face=Arial color=#00529B><%=FormatNumber(RS1.Fields("saldohoy"),2)%></font><%if DivisaDif<>"" then%><table cellspacing=0 cellpadding=0 border=0 width="100%"><tr><td width="50%" align="center"><font size=1 face=Arial color=#00529B><%=RS1.Fields("Divisa")%></font></td><td width="50%" align="center"><font size=1 face=Arial color=#00529B><%=DivisaDif%></font></td></tr><tr><td align="center"><font size=1 face=Arial color=#00529B><%=FormatNumber(MontoTotalDivisa,2)%></font></td><td align="center"><font size=1 face=Arial color=#00529B><%=FormatNumber(MontoTotalDivisaDif,2)%></font></td></tr></table><%end if%></td>
			</tr>
				<%
				if nrocuotas>0 then
				%>
					<tr style="display: none" id="<%=RS1.Fields("Contrato")%>">
							<td valign="top" align="center" colspan=7>
								<table cellspacing=2 cellpadding=2 border=0 width="95%">
									<tr bgcolor="#BEE8FB">
										<td align="center" rowspan=2 width="1%"><font size=1 face=Arial color=#00529B>N°</font></td>
										<td align="center" rowspan=2 width="4%"><font size=1 face=Arial color=#00529B>F.Venc.</font></td>
										<td align="center" width="19%" <%if DivisaDif<>"" then%> colspan="2"<%end if%>><font size=1 face=Arial color=#00529B>Capital</font></td>
										<td align="center" width="19%" <%if DivisaDif<>"" then%> colspan="2"<%end if%>><font size=1 face=Arial color=#00529B>Interés</font></td>
										<td align="center" width="19%" <%if DivisaDif<>"" then%> colspan="2"<%end if%>><font size=1 face=Arial color=#00529B>Gast/Com/Otros</font></td>
										<td align="center" width="19%" <%if DivisaDif<>"" then%> colspan="2"<%end if%>><font size=1 face=Arial color=#00529B>Int.Venc/Mor</font></td>
										<td align="center" width="19%" <%if DivisaDif<>"" then%> colspan="2"<%end if%>><font size=1 face=Arial color=#00529B>Total</font></td>
									</tr>
									<tr bgcolor="#BEE8FB">
										<td align="center"><font size=1 face=Arial color=#00529B><%=RS1.Fields("Divisa")%></font></td>
										<%if DivisaDif<>"" then%><td align="center" width="9.5%"><font size=1 face=Arial color=#00529B><%=DivisaDif%></font></td><%end if%>
										<td align="center"><font size=1 face=Arial color=#00529B><%=RS1.Fields("Divisa")%></font></td>
										<%if DivisaDif<>"" then%><td align="center" width="9.5%"><font size=1 face=Arial color=#00529B><%=DivisaDif%></font></td><%end if%>
										<td align="center"><font size=1 face=Arial color=#00529B><%=RS1.Fields("Divisa")%></font></td>
										<%if DivisaDif<>"" then%><td align="center" width="9.5%"><font size=1 face=Arial color=#00529B><%=DivisaDif%></font></td><%end if%>
										<td align="center"><font size=1 face=Arial color=#00529B><%=RS1.Fields("Divisa")%></font></td>
										<%if DivisaDif<>"" then%><td align="center" width="9.5%"><font size=1 face=Arial color=#00529B><%=DivisaDif%></font></td><%end if%>
										<td align="center"><font size=1 face=Arial color=#00529B><%=RS1.Fields("Divisa")%></font></td>
										<%if DivisaDif<>"" then%><td align="center" width="9.5%"><font size=1 face=Arial color=#00529B><%=DivisaDif%></font></td><%end if%>
									</tr>									
									<%
									contador=0
									TotalCapital1=0
									TotalCapital2=0
									TotalInteres1=0
									TotalInteres2=0
									TotalComision1=0
									TotalComision2=0
									TotalInteresMora1=0
									TotalInteresMora2=0																		
									RS2.MoveFirst
									Do While not RS2.EOF
										contador=contador + 1
										fechavencimiento=RS2.Fields("FechaVencimiento")
																				
										Capital1=0
										Interes1=0
										Comision1=0
										InteresMora1=0
										Capital2=0
										Interes2=0
										Comision2=0
										InteresMora2=0
											
										if RS2.Fields("Divisa")<>DivisaDif then
											Capital1=RS2.Fields("Capital")
											Interes1=RS2.Fields("Interes")
											if RS1.Fields("producto")="50" then
											Comision1=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("Seguro") + RS2.Fields("InteresMoratorio")
											else
											Comision1=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("InteresMoratorio")
											end if	
											InteresMora1=RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
																				
											TotalCapital1=TotalCapital1 + Capital1
											TotalInteres1=TotalInteres1 + Interes1
											TotalComision1=TotalComision1 + Comision1
											TotalInteresMora1=TotalInteresMora1 + InteresMora1
										else
											Capital2=RS2.Fields("Capital")
											Interes2=RS2.Fields("Interes")
											if RS1.Fields("producto")="50" then
											Comision2=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("Seguro") + RS2.Fields("InteresMoratorio")
											else
											Comision2=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("InteresMoratorio")
											end if	
											InteresMora2=RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
																					
											TotalCapital2=TotalCapital2 + Capital2
											TotalInteres2=TotalInteres2 + Interes2
											TotalComision2=TotalComision2 + Comision2
											TotalInteresMora2=TotalInteresMora2 + InteresMora2
										end if
										
										if not RS2.EOF then RS2.MoveNext
										
										if not RS2.EOF then 										
											if fechavencimiento=RS2.Fields("fechavencimiento") then
												if RS2.Fields("Divisa")<>DivisaDif then
													Capital1=RS2.Fields("Capital")
													Interes1=RS2.Fields("Interes")
													if RS1.Fields("producto")="50" then
													Comision1=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("Seguro") + RS2.Fields("InteresMoratorio")
													else
													Comision1=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("InteresMoratorio")
													end if	
													InteresMora1=RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
																						
													TotalCapital1=TotalCapital1 + Capital1
													TotalInteres1=TotalInteres1 + Interes1
													TotalComision1=TotalComision1 + Comision1
													TotalInteresMora1=TotalInteresMora1 + InteresMora1
												else
													Capital2=RS2.Fields("Capital")
													Interes2=RS2.Fields("Interes")
													if RS1.Fields("producto")="50" then
													Comision2=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("Seguro") + RS2.Fields("InteresMoratorio")
													else
													Comision2=RS2.Fields("Comision") + RS2.Fields("Gasto") + RS2.Fields("InteresMoratorio")
													end if	
													InteresMora2=RS2.Fields("interesdemora") + RS2.Fields("interesvencido")
																							
													TotalCapital2=TotalCapital2 + Capital2
													TotalInteres2=TotalInteres2 + Interes2
													TotalComision2=TotalComision2 + Comision2
													TotalInteresMora2=TotalInteresMora2 + InteresMora2
												end if										
												RS2.MoveNext
											end if
										end if
									%>
									<tr bgcolor="#E9F8FE">
										<td align="center"><font size=1 face=Arial color=#00529B><%=contador%></font></td>
										<td align="center"><font size=1 face=Arial color=#00529B><%=fechavencimiento%></font></td>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Capital1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Capital2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Interes1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Interes2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Comision1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Comision2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(InteresMora1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(InteresMora2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Capital1 + Interes1 + Comision1 + InteresMora1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(Capital2 + Interes2 + Comision2 + InteresMora2,2)%></font></td><%end if%>
									</tr>
									<%
									Loop
									%>
									<tr bgcolor="#BEE8FB">
										<td align="center"><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
										<td align="center"><font size=1 face=Arial color=#00529B>Total</font></td>
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalCapital1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalCapital2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalInteres1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalInteres2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalComision1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalComision2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalInteresMora1,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(TotalInteresMora2,2)%></font></td><%end if%>
										
										<td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(MontoTotalDivisa,2)%></font></td>
										<%if DivisaDif<>"" then%><td align="right"><font size=1 face=Arial color=#00529B><%=FormatNumber(MontoTotalDivisaDif,2)%></font></td><%end if%>
									</tr>									
								</table>
							</td>
					</tr>							
				<%
				end if
				%>
			<%
			RS2.Close
			RS1.MoveNext
			Loop
			RS1.Close
			%>
			</table>
			
		<%if contador=0 then%>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Gestiones (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
			</tr>
			</table>
		<%else%>
			<!--<table>
			<tr bgcolor="#E9F8FE">
			    
			<tr>
			</table>		--->
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Gestiones (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:agregargestion();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
                <td bgcolor="#F5F5F5" align="right"><font size=2 face=Arial color=#00529B><b>Fecha inicio:</td>
			    <td bgcolor="#F5F5F5"><input name="fechagestionini" type=text maxlength=10 id="sel1" value="<%if IsDate(fechagestionini) then%><%=fechagestionini%><%end if%>" style="font-size: x-small; width: 70px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel1', '%d/%m/%Y');">
			    <td bgcolor="#F5F5F5" align="right"><font size=2 face=Arial color=#00529B><b>Fecha fin:</td>
			    <td bgcolor="#F5F5F5"><input name="fechagestionfin" type=text maxlength=10 id="sel2" value="<%if IsDate(fechagestionfin) then%><%=fechagestionfin%><%end if%>" style="font-size: x-small; width: 70px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel2', '%d/%m/%Y');">
			    <td bgcolor="#F5F5F5"><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>				
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
			</tr>
			</table>
			<div id="tabla0"> 
			</div>
		<%end if%>
			<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
			<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
			<input type="hidden" name="codigocentral" value="<%=codigocentral%>">
			<input type="hidden" name="contrato" value="<%=contrato%>">
			<input type="hidden" name="fechadatos" value="<%=fechadatos%>">
			<input type="hidden" name="fechagestion" value="<%=fechagestion%>">
			<input type="hidden" name="direccioneliminar" value="">
			<input type="hidden" name="telefonoeliminar" value="">
			<input type="hidden" name="emaileliminar" value="">		
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
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>							
		</body>
		</html>		
		<%
		''Codigo exp excel
		''Si se pide exportar a excel
				if expimp="1" then
					''Para Exportar a Excel
					''Paso Cero eliminar exportación anterior
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"set @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Primero Cabecera en temp1_(user).txt
					consulta_exp="select 'Fecha de Gestión','Agencia','Gestion','Comentario','F.Promesa','Direccion/ Telefono','Adjunto'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select convert(varchar,A.fhgestionado,103) + ' ' + convert(varchar,A.fhgestionado,108),C.RazonSocial,B.Descripcion, A.comentario,ISNULL(convert(varchar,A.fechapromesa,103),''),CASE WHEN ltrim(rtrim(A.fono))<>'' THEN rtrim(ltrim(A.tipofono)) + '-' + rtrim(ltrim(A.prefijo)) + '-' + rtrim(ltrim(A.fono)) + (CASE WHEN LEN(RTRIM(ltrim(A.extension)))> 0 and A.extension<>'0000' THEN '-'+ RTRIM(LTRIM(A.extension)) ELSE '' END) ELSE A.direccion END,CASE WHEN A.ficherogestion IS NOT NULL THEN 'Sí' ELSE 'No' END " & _
								 "from CobranzaCM.dbo.RespuestaGestion A inner join CobranzaCM.dbo.Gestion B on A.codgestion=B.codgestion left outer join CobranzaCM.dbo.Agencia C on A.codagencia=C.codagencia " & filtrobuscador & " order by A.fhgestionado desc"
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
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>



