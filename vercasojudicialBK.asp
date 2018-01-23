<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repsaeestudio.asp") or permisofacultad("repsaeterritorio.asp") or permisofacultad("repsaegestor.asp") then
	
		
		 ''codigocentral=obtener("codigocentral")
		contrato=obtener("contrato")
		''fechadatos=obtener("fechadatos")
		''buscador=obtener("buscador")
		fechacierre=mid(obtener("fechacierre"),7,4) & mid(obtener("fechacierre"),4,2) & mid(obtener("fechacierre"),1,2)
		''fechagestion=obtener("fechagestion")
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
		''if not IsDate(fechagestionini) then
		''    fechagestionini=CStr(minfechagestion)
		''end if
		''if not IsDate(fechagestionfin) then
		''    fechagestionfin=CStr(maxfechagestion)
		''end if	
		
		sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload' or descripcion='RutaWebUpload'"
		consultar sql,RS3
		RS3.Filter=" descripcion='RutaFisicaUpload'"
		RutaFisicaUpload=RS3.Fields(1)
		RS3.Filter=" descripcion='RutaWebUpload'"
		RutaWebUpload=RS3.Fields(1)				
		RS3.Filter=""
		RS3.Close	
		
				
		
		''if mid(obtener("fechagestion"),7,4) & mid(obtener("fechagestion"),4,2) & mid(obtener("fechagestion"),1,2)=CStr(maxfechagestion) then
		''	vistabusqueda="VERULTIMPAGADO"
		''else
		''	vistabusqueda="VERIMPAGADO"
		''end if				
		
		

		
		'''sql="select diasvencimiento,nombre,agencia,tipodocumento,numdocumento,direccion,departamento,provincia,distrito,tipofono1,prefijo1,fono1,extension1,tipofono2,prefijo2,fono2,extension2,tipofono3,prefijo3,fono3,extension3,tipofono4,prefijo4,fono4,extension4,tipofono5,prefijo5,fono5,extension5,email,(select top 1 Y.descripcion from MarcaCliente X inner join Marca Y on X.codigocentral=A.codigocentral and X.codmarca=Y.codmarca and X.activo=1 order by X.codmarcacliente desc) as Marca from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & fechadatos & "' order by DiasVencimiento desc"
		''sql="select *,(select top 1 Y.descripcion from MarcaCliente X inner join Marca Y on X.codigocentral=A.codigocentral and X.codmarca=Y.codmarca and X.activo=1 order by X.codmarcacliente desc) as Marca,(select count(distinct fechavencimiento) from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos) as NumCuotas,(select top 1 divisa from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos and divisa<>A.divisa) as DivisaDif from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & fechadatos & "' order by DiasVencimiento desc,saldohoy desc"
		''sql="select *,(select top 1 Y.descripcion from MarcaCliente X inner join Marca Y on X.codigocentral=A.codigocentral and X.codmarca=Y.codmarca and X.activo=1 order by X.codmarcacliente desc) as Marca from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' order by DiasVencimiento desc,saldohoy desc"
		''sql="select * from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(obtener("fechadatos"),7,4) & mid(obtener("fechadatos"),4,2) & mid(obtener("fechadatos"),1,2) & "' order by DiasVencimiento desc,saldohoy desc"
		

		
		sql="select A.CONTRATO,A.CODCENTRAL,A.CLIENTE,A.PRODUCTO,A.DA,A.FLAG_REF,A.CLASIFICACION,A.JUD,A.PROVCONST, " &_
		    "A.GARANT_TOTAL,A.CODOFIC,A.CODTERRITORIO,A.TERRITORIO,A.OFICINA,A.FECHAPASEAMORA,A.SEGMENTO_RIESGO,A.Plaza,A.ESTUDIO,A.ESPECIALISTA,A.FASIGNA,A.PROCESO, " &_
		    "A.FECFASE,A.FASE,(CASE WHEN A.DNI<>'' THEN 'DNI - ' + A.DNI ELSE 'RUC - '+ A.RUC END) AS DOCUMENTO,(select top 1 fechaformal from FB.CentroCobranzas.dbo.PD_Detalle_Casos_JUD where contrato=A.contrato and fechadatos=A.fechadatos) as FechaFormal, B.DESCRIPCION_SGTE, " &_
		    " CASE WHEN (select count(*) from FB.CentroCobranzas.dbo.PD_Detalle_Casos_Jud X inner join FB.CentroCobranzas.dbo.Fases_SAE Y on X.CODIGO=Y.CODIGO and X.TIPOPROCESO=Y.TIPOPROCESO where X.contrato=A.contrato and X.fechadatos=A.fechadatos and Y.Embargo=1)>1 THEN 'SI' ELSE 'NO' END AS EMBARGO, " &_
		    "(select min(X.FECFASE) from FB.CentroCobranzas.dbo.PD_Detalle_Casos_Jud X inner join FB.CentroCobranzas.dbo.Fases_SAE Y on X.CODIGO=Y.CODIGO and X.TIPOPROCESO=Y.TIPOPROCESO where X.contrato=A.contrato and X.fechadatos=A.fechadatos and Y.Demanda=1) as FechaDemanda " &_
		    "from FB.CentroCobranzas.dbo.PD_Casos_Jud A LEFT OUTER JOIN FB.CentroCobranzas.dbo.FASES_SAE B On A.CODFASE=B.CODIGO and A.TIPOPROCESO=B.TIPOPROCESO where A.CONTRATO='" & contrato & "' and A.FECHADATOS='" & fechacierre & "'"
		''Response.Write sql
		consultar sql,RS2
		if not RS2.EOF then
			codcentral=RS2.Fields("CODCENTRAL")
			cliente=RS2.Fields("CLIENTE")
			documento=RS2.Fields("DOCUMENTO")
			producto=RS2.Fields("PRODUCTO")
			diasvencimiento=RS2.Fields("DA")
			refinanciado=RS2.Fields("FLAG_REF")
			clasificacion=RS2.Fields("CLASIFICACION")
			judicial=RS2.Fields("JUD")
			provconst=RS2.Fields("PROVCONST")
	        garantia=RS2.Fields("GARANT_TOTAL")
	        codterritorio=RS2.Fields("CODTERRITORIO")
	        codoficina=RS2.Fields("CODOFIC")
	        territorio=RS2.Fields("TERRITORIO")
	        oficina=RS2.Fields("OFICINA")
	        fechapaseamora=RS2.Fields("FECHAPASEAMORA")
	        segmento_riesgo=RS2.Fields("SEGMENTO_RIESGO")
	        plaza=RS2.Fields("Plaza")
	        estudio=RS2.Fields("ESTUDIO")
			especialista=RS2.Fields("ESPECIALISTA")
			fasigna=RS2.Fields("FASIGNA")
			proceso=RS2.Fields("PROCESO")
			fecfase=RS2.Fields("FECFASE")
			fase=RS2.Fields("FASE")
			fechaformal=RS2.Fields("FechaFormal")
			embargo=RS2.Fields("EMBARGO")
			FechaDemanda=RS2.Fields("FechaDemanda")
			Fasesiguiente=RS2.Fields("DESCRIPCION_SGTE")
		end if
		RS2.Close
		

		sql4="select top 1 A.GARANT_RR, A.GARANT_PREF, A.GARANT_AUTO, A.GARANT_CONTRA, /*A.JUZGADO,*/ (select PROVINCIA + ' / ' + CIUDAD from FB.CentroCobranzas.dbo.SAEJZGDO_CODIGO_CIUDAD where CLAVE=(select MAX(RTRIM(LTRIM(ESTADOEXP)) + LTRIM(RTRIM(CIUDAD))) FROM FB.CentroCobranzas.dbo.XCOM_SAEAEDTPCT WHERE CONTRATO=A.CONTRATO and FECHADATOS=A.FECHADATOS)) as JUZGADO, A.NROEXPJUDI  " &_
		    "from FB.CentroCobranzas.dbo.PD_Detalle_Casos_Jud A where A.CONTRATO='" & contrato & "' and A.FECHADATOS='" & fechacierre & "' ORDER BY A.FECHADATOS DESC"
		''Response.Write sql
		consultar sql4,RS3
		if not RS3.EOF then
			garantpref=RS3.Fields("GARANT_PREF")
			garantrr=RS3.Fields("GARANT_RR")
			garantauto=RS3.Fields("GARANT_AUTO")
			garantcontra=RS3.Fields("GARANT_CONTRA")
			juzgado=RS3.Fields("JUZGADO")
			expediente=RS3.Fields("NROEXPJUDI")
		end if
		RS3.Close
		
		sql4="select rtrim(ltrim(obs)) as Observaciones,IsDate(SUBSTRING(ltrim(rtrim(obs)),1,10)) as Fecha " &_
		    "from FB.CentroCobranzas.dbo.XCOM_AE_Observaciones A where A.CONTRATO='" & contrato & "' and A.FECHADATOS='" & fechacierre & "' order by FechaObs"
		''Response.Write sql
		observaciones=""
		consultar sql4,RS3
		do while not RS3.EOF
		    if RS3.Fields("Fecha")=1 then
		        if observaciones="" then
		        observaciones=observaciones & RS3.Fields("Observaciones")
		        else
		        observaciones=observaciones & "<BR>" &  RS3.Fields("Observaciones")
		        end if
		    else
			observaciones=observaciones & RS3.Fields("Observaciones")
			end if
		RS3.MoveNext
		Loop
		RS3.Close		
		
		sql5="select sum(A.JUD) as DEUDATOTAL " &_
		    "from FB.CentroCobranzas.dbo.PD_Detalle_Casos_Jud A where A.codcentral='" & codcentral & "' and A.FECHADATOS='" & fechacierre & "'"
		''Response.Write sql
		consultar sql5,RS3
		if not RS3.EOF then
			deudatotal=RS3.Fields("DEUDATOTAL")
		end if
		RS3.Close

		sql2="select A.* from FB.CentroCobranzas.dbo.maeclien A inner join FB.CentroCobranzas.dbo.PD_Casos_JUD B on A.codcent=B.CODCENTRAL where B.CONTRATO='" & contrato & "'" 
		consultar sql2,RS3	
		if not RS3.EOF then
		
			email=RS3.Fields("email")
			direccion=RS3.Fields("dirprin")
			distrito=RS3.Fields("desdistr")
			provincia=RS3.Fields("desprovi")
			departamento=RS3.Fields("desdpto")
			tipofono1=RS3.Fields("tiptel1")
			prefijo1=RS3.Fields("pretel1")
			fono1=RS3.Fields("numtel1")
			extension1=RS3.Fields("exttel1")
			tipofono2=RS3.Fields("tiptel2")
			prefijo2=RS3.Fields("pretel2")
			fono2=RS3.Fields("numtel2")
			extension2=RS3.Fields("exttel2")
			tipofono3=RS3.Fields("tiptel3")
			prefijo3=RS3.Fields("pretel3")
			fono3=RS3.Fields("numtel3")
			extension3=RS3.Fields("exttel3")
			
		end if
		RS3.Close
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<title>Ver Caso Judicial</title>
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
			nuevotelf=global_popup_IWTSystem(nuevotelf,"adicionartelefono.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newtelf","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
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
			nuevoemail=global_popup_IWTSystem(nuevoemail,"adicionaremail.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","Newemail","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=170,width=" + (screen.width/3 - 100) + ",left=" + (screen.width/4 + 150 ) + ",resizable=yes");
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
			nuevogestion=global_popup_IWTSystem(nuevogestion,"adicionarseguimiento.asp?vistapadre1=" + window.name + "&paginapadre1=verimpagado.asp&vistapadre=" + window.name + "&paginapadre=<%=obtener("paginapadre")%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","NewGestion","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 250)/2 - 30) + ",height=300,width=" + (screen.width/2 + 100) + ",left=" + (screen.width/4 - 50) + ",resizable=yes");
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
		
		''filtrobuscador = "where A.fhgestionado between convert(datetime,'" & mid(fechagestionini,7,4) & mid(fechagestionini,4,2) & mid(fechagestionini,1,2) & "') and convert(datetime,'" & mid(fechagestionfin ,7,4) & mid(fechagestionfin,4,2) & mid(fechagestionfin,1,2) & "') and A.codigocentral='" & codigocentral & "'"
		filtrobuscador = "where A.contrato='" & contrato & "'"
				
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
		    sql="select top " & cantidadxpagina & " A.codrespgestion,A.fhgestionado,A.contrato,A.comentario,A.fechapromesa,A.fono,A.tipofono,A.prefijo,A.extension,A.direccion,A.ficherogestion,C.RazonSocial,B.Descripcion,UB.distrito,UB.provincia,UB.departamento from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia LEFT OUTER JOIN Ubigeo AS UB ON UB.coddpto = A.coddpto AND UB.codprov = A.codprov AND UB.coddist = A.coddist where " & filtrobuscador1 & " A.codrespgestion not in (select top " & topnovisible & " A.codrespgestion from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia " & filtrobuscador & " order by A.fhgestionado desc) order by A.codrespgestion desc" 
		else
		    sql="select top " & cantidadxpagina & " A.codrespgestion,A.fhgestionado,A.contrato,A.comentario,A.fechapromesa,A.fono,A.tipofono,A.prefijo,A.extension,A.direccion,A.ficherogestion,C.RazonSocial,B.Descripcion,UB.distrito,UB.provincia,UB.departamento from RespuestaGestion A inner join Gestion B on A.codgestion=B.codgestion left outer join Agencia C on A.codagencia=C.codagencia LEFT OUTER JOIN Ubigeo AS UB ON UB.coddpto = A.coddpto AND UB.codprov = A.codprov AND UB.coddist = A.coddist " & filtrobuscador & " order by A.codrespgestion desc"
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
				datos[tabla][<%=contador%>][5]='<%if len(trim(RS.Fields("comentario")))<=25 then%><%=trim(replace(replace(RS.Fields("comentario"),chr(10),""),chr(13),""))%><%else%><%=mid(trim(replace(replace(RS.Fields("comentario"),chr(10),""),chr(13),"")),1,25) & "..."%><%end if%>';
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
				<td align="center" height="22"><font size=2 face=Arial color="#FFFFFF"><b>Módulo Gestión de Casos Judiciales</b></font></td>
			</tr>
			</table>
			<table width=100% cellpadding=2 cellspacing=2 border=0>

			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Código Central:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codcentral%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Nombres:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=cliente%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Tipo y N° Doc:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=documento%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Segmento:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=segmento_riesgo%></b></font></td>
			</tr>
            <tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Oficina:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codoficina%>&nbsp;-&nbsp;<%=oficina%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Territorio:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=codterritorio%>&nbsp;-&nbsp;<%=territorio%></b></font></td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Fecha Asignación:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=fasigna%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Estudio:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=estudio%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Especialista:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=especialista%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Plaza:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=plaza%></b></font></td>
			</tr>						
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Contrato:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=contrato%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Producto:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=producto%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Fecha Form.:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=fechaformal%></b></font></td>			
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Fecha pase a mora:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=fechapaseamora%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Deuda Total S/.:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=formatnumber(deudatotal,2)%></b></font></td>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Juzgado/Expediente:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=juzgado%>&nbsp;/&nbsp;<%=expediente%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Deuda Contrato S/.:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=formatnumber(judicial,2)%></b></font></td>	
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Clasificación:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=clasificacion%></b></font></td>			 						 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" width=15%><font size=2 face=Arial color=#00529B>&nbsp;Prov. Const:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=formatnumber(provconst,2)%></b></font></td>				
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Días Vencimiento:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=diasvencimiento%></b></font></td>
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Embargo:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=embargo%></b></font></td>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Fecha Demanda:</font></td>
				 <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;<%=FechaDemanda%></b></font></td>				 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Fase Siguiente:</font></td>
				 <td bgcolor="#E9F8FE" colspan=3><font size=2 face=Arial color=#00529B><b>&nbsp;<%=Fasesiguiente%></b></font></td>			 
			</tr>
			<tr>
				 <td bgcolor="#BEE8FB" valign=top><font size=2 face=Arial color=#00529B>&nbsp;Observaciones:</font></td>
				 <td bgcolor="#E9F8FE" colspan=3><font size=2 face=Arial color=#00529B><%=Observaciones%></font></td>			 
			</tr>			
			<!--<tr>
			
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Garantías:</font></td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>-->
						<script language=javascript>
						    function visualizargarantia() {
						        var filas = document.getElementById('tablagarantias').rows.length;
						        //if (document.getElementById('imagengarantia').src.length==document.getElementById('imagengarantia').src.replace('mostrar','').length)
						        if (document.getElementById('imagengarantia').title == "Mostrar") {
						            document.getElementById('imagengarantia').title = "Ocultar";
						            document.getElementById('imagengarantia').alt = "Ocultar";
						            document.getElementById('imagengarantia').src = "imagenes/ocultar.png";
						            for (i = 1; i < filas; i++) {
						                document.getElementById('tablagarantias').rows[i].style.display = '';
						            }
						        }
						        else {
						            document.getElementById('imagengarantia').title = "Mostrar";
						            document.getElementById('imagengarantia').alt = "Mostrar";
						            document.getElementById('imagengarantia').src = "imagenes/mostrar.png";
						            for (i = 1; i < filas; i++) {
						                document.getElementById('tablagarantias').rows[i].style.display = 'none';
						            }
						        }
						    }
						</script>
				 		<!--<table id="tablagarantias" width=100% cellpadding=0 cellspacing=1 border=0>
						<tr>
							<td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizargarantia();"><img id="imagengarantia" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
							<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>&nbsp;Garantía Total:</b></font></td>
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><b>&nbsp;S/.<%=formatnumber(garantia,2)%></b></font></td>
						</tr>
						<% if garantrr<>0 then %>			 		
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td  width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía RR:</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantrr,2)%></font></td>

							</tr>
						<% end if %>
						<% if garantpref<>0 then %>	
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía Preferida (Hipotecaria):</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantpref,2)%></font></td>
							</tr>
						<% end if %>
						
						<% if garantauto<>0 then %>
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía Autoliquidable:</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantauto,2)%></font></td>
							</tr>
						<% end if %>
						<% if garantcontra<>0 then %>
							<tr style="display: none">
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
								<td width=250 bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B>&nbsp;Garantía Contraparte:</font></td>
								<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;S/.<%=formatnumber(garantcontra,2)%></font></td>
							</tr>
						<% end if %>
						</table>
				 </td>
			</tr>-->
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Hipotecas:</font></td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
                        <%
                            sql="select " &_
		                            "A.C_CENT_CLIE as CodigoCentral, " &_
		                            "A.N_Garantia as NroGarantia, " &_
		                            "B.I_RANGO as Rango, " &_
		                            "ltrim(rtrim(b.D_DIRE_PROPIEDAD_TEXT)) as Direccion, " &_
		                            "ltrim(rtrim(b.D_NOMB_OTOR_TEXT)) as Otorgante,   " &_
		                            "A.c_moneda as Moneda, " &_
		                            "B.M_TASADO as VComercial,  " &_
		                            "B.M_REAL as VRealizacion, " &_
		                            "A.M_GARANTIA as VGravamen, " &_
		                            "case " &_
		                            "when b.i_tipo_bien=1 then 'Casa Habitación'  " &_
		                            "when b.i_tipo_bien=2 then 'Local Comercial'  " &_
		                            "when b.i_tipo_bien=3 then 'Local Industrial'  " &_
		                            "when b.i_tipo_bien=4 then 'Terreno'  " &_
		                            "when b.i_tipo_bien=5 then 'Otros'  " &_
		                            "else '' end as bien,		 " &_
		                            "CASE WHEN A.I_TIPO_GARA ='111' THEN 'Hipoteca' else '' end AS TIPOGARANTIA, " &_
		                            "A.FECHADATOS " &_
                            "from FB.CentroCobranzas.dbo.XCOM_CRDT000_Datos_Generales A " &_
                            "inner join FB.CentroCobranzas.dbo.XCOM_CRDT001_HIPOTECAS B " &_
                            "on A.I_TIPO_GARA IN ('111')  " &_
                            "AND A.C_CENT_CLIE='" & codcentral & "'  " &_
                            "AND A.I_SITUACION IN ('6') " &_
                            "and B.N_detalle=1 " &_
                            "and A.N_GARANTIA= B.N_GARANTIA " &_
                            "and A.FECHADATOS=B.FECHADATOS " &_
                            "where A.FECHADATOS=(select MAX(FECHADATOS) from FB.CentroCobranzas.dbo.XCOM_CRDT001_HIPOTECAS)"
                            
                            consultar sql,RS1
                            if RS1.RecordCount>0 then
                                %>  
								            <table width=100% cellpadding=0 cellspacing=2 border=0>
								        	    <tr>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>N°</b></font></td>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Dirección</b></font></td>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Otorgante</b></font></td>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Bien</b></font></td>
                                                    <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Moneda</b></font></td>
                                                    <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>V.Comercial</b></font></td>
                                                    <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>V.Realización</b></font></td>
                                                    <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>V.Gravamen</b></font></td>
								        	    </tr>
								        	    <%do while not RS1.EOF%>
								        	    <tr>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("NroGarantia")%></font></td>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Direccion")%></font></td>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Otorgante")%></font></td>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Bien")%></font></td>
                                                    <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Moneda")%></font></td>
                                                    <td bgcolor="#E9F8FE" align=right><font size=2 face=Arial color=#00529B><%=FormatNumber(RS1.Fields("VComercial"),2)%></font></td>
                                                    <td bgcolor="#E9F8FE" align=right><font size=2 face=Arial color=#00529B><%=FormatNumber(RS1.Fields("VRealizacion"),2)%></font></td>
                                                    <td bgcolor="#E9F8FE" align=right><font size=2 face=Arial color=#00529B><%=FormatNumber(RS1.Fields("VGravamen"),2)%></font></td>
								        	    </tr>				
								        	    <%
								        	    RS1.movenext
								        	    loop%>				        	    
    						                </table>
						        <%
						    else
						    %>
					            <table width=100% cellpadding=0 cellspacing=2 border=0>
					        	    <tr>
					        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>No cuenta con Hipotecas.</font></td>						    
					        	    </tr>
					            </table>                    
						    <%
						    end if 
						   RS1.Close
						   %>
				 </td>
			</tr>	   
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Fiadores:</font></td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
                        <%
                            
                            sql="select distinct C.CodCent as CodigoCentral, " &_
		                                    "RTRIM(LTRIM(ISNULL(C.nombre,''))) + ' ' + RTRIM(LTRIM(ISNULL(C.apellido1,''))) + ' ' + RTRIM(LTRIM(ISNULL(C.apellido2,''))) as Nombres, " &_
		                                    "C.pretel1 + ' - ' + C.numtel1 + CASE WHEN C.exttel1<>'0000' and C.exttel1<>'' THEN ' - ' + C.exttel1 ELSE '' END as Telefono, " &_
		                                    "CASE WHEN A.I_TIPO_GARA ='311' THEN 'Fiador' ELSE '' end AS TIPOGARANTIA,  " &_
		                                    "A.FECHADATOS  " &_
                                    "from FB.CentroCobranzas.dbo.XCOM_CRDT000_Datos_Generales A  " &_
                                    "inner join FB.CentroCobranzas.dbo.XCOM_CRDT014_Fianza_Solidaria B  " &_
                                    "on A.I_TIPO_GARA IN ('311')  " &_
                                    "AND A.C_CENT_CLIE='" & codcentral & "'  " &_
                                    "AND A.I_SITUACION IN ('6')  " &_
                                    "and B.detalle=1  " &_
                                    "and A.N_GARANTIA=B.GARANTIA  " &_
                                    "and A.FECHADATOS=B.FECHADATOS  " &_
                                    "left outer join FB.CentroCobranzas.dbo.maeclien C " &_
                                    "on B.FIADOR=C.CodCent " &_
                                    "where A.FECHADATOS=(select MAX(FECHADATOS) from FB.CentroCobranzas.dbo.XCOM_CRDT014_Fianza_Solidaria)"
                            consultar sql,RS1
                            if RS1.RecordCount>0 then
                                %>  
								            <table width=100% cellpadding=0 cellspacing=2 border=0>
								        	    <tr>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Cód.Central</b></font></td>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Nombres</b></font></td>
								        	        <td bgcolor="#BEE8FB"><font size=2 face=Arial color=#00529B><b>Teléfono</b></font></td>
								        	    </tr>
								        	    <%do while not RS1.EOF%>
								        	    <tr>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("CodigoCentral")%></font></td>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("Nombres")%></font></td>
								        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B><%=RS1.Fields("telefono")%></font></td>
								        	    </tr>				
								        	    <%
								        	    RS1.movenext
								        	    loop%>				        	    
    						                </table>
						     <%
						    else
						    %>
					            <table width=100% cellpadding=0 cellspacing=2 border=0>
					        	    <tr>
					        	        <td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>No cuenta con Fiadores.</font></td>						    
					        	    </tr>
					            </table>                    
						    <%
						    end if 
						   RS1.Close
						   %>
				 </td>
			</tr>	 			         
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Direcciones:</font></td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
						<script language=javascript>
						    function visualizardir() {
						        var filas = document.getElementById('tabladirecciones').rows.length;
						        //if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
						        if (document.getElementById('imagendir').title == "Mostrar") {
						            document.getElementById('imagendir').title = "Ocultar";
						            document.getElementById('imagendir').alt = "Ocultar";
						            document.getElementById('imagendir').src = "imagenes/ocultar.png";
						            for (i = 2; i < filas; i++) {
						                document.getElementById('tabladirecciones').rows[i].style.display = '';
						            }
						        }
						        else {
						            document.getElementById('imagendir').title = "Mostrar";
						            document.getElementById('imagendir').alt = "Mostrar";
						            document.getElementById('imagendir').src = "imagenes/mostrar.png";
						            for (i = 2; i < filas; i++) {
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
						
						sql="select A.coddireccionnueva,A.direccion,B.departamento,B.provincia,B.distrito from DireccionNueva A left outer join Ubigeo B on A.coddpto=B.coddpto and A.codprov=B.codprov and A.coddist=B.coddist inner join FB.CentroCobranzas.dbo.PD_Casos_JUD C on A.codigocentral=C.CODCENTRAL where C.CONTRATO='" & contrato & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.coddireccionnueva desc"
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
				 </td>
			</tr>	
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;Teléfonos:</font></td>
					</tr>
					</table>						
				 </td>
				 <td colspan=3>
						<script language=javascript>
						    function visualizartelf() {
						        var filas = document.getElementById('tablatelefonos').rows.length;
						        if (document.getElementById('imagentel').title == "Mostrar") {
						            document.getElementById('imagentel').title = "Ocultar";
						            document.getElementById('imagentel').alt = "Ocultar";
						            document.getElementById('imagentel').src = "imagenes/ocultar.png";
						            for (i = 2; i < filas; i++) {
						                document.getElementById('tablatelefonos').rows[i].style.display = '';
						            }
						        }
						        else {
						            document.getElementById('imagentel').title = "Mostrar";
						            document.getElementById('imagentel').alt = "Mostrar";
						            document.getElementById('imagentel').src = "imagenes/mostrar.png";
						            for (i = 2; i < filas; i++) {
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
						<%	''mae clien	
						sql="select codtelefononuevo,codtipotelefono,prefijo,fono,extension from TelefonoNuevo A inner join FB.CentroCobranzas.dbo.PD_Casos_JUD C on A.codigocentral=C.CODCENTRAL where C.CONTRATO='" & contrato & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codtelefononuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
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
				 </td>
			</tr>							
			<tr>
				 <td bgcolor="#BEE8FB" valign=top>
					<table width=100% cellpadding=0 cellspacing=0 border=0>
					<tr>
						<td><font size=2 face=Arial color=#00529B>&nbsp;E-mail:</font></td>
					</tr>
					</table>					
				 </td>
				 <td colspan=3>
						<script language=javascript>
						    function visualizaremail() {
						        var filas = document.getElementById('tablaemails').rows.length;
						        //if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
						        if (document.getElementById('imagenemail').title == "Mostrar") {
						            document.getElementById('imagenemail').title = "Ocultar";
						            document.getElementById('imagenemail').alt = "Ocultar";
						            document.getElementById('imagenemail').src = "imagenes/ocultar.png";
						            for (i = 2; i < filas; i++) {
						                document.getElementById('tablaemails').rows[i].style.display = '';
						            }
						        }
						        else {
						            document.getElementById('imagenemail').title = "Mostrar";
						            document.getElementById('imagenemail').alt = "Mostrar";
						            document.getElementById('imagenemail').src = "imagenes/mostrar.png";
						            for (i = 2; i < filas; i++) {
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
						sql="select A.codemailnuevo,A.email from EmailNuevo A inner join FB.CentroCobranzas.dbo.PD_Casos_JUD C on A.codigocentral=C.CODCENTRAL where C.CONTRATO='" & contrato & "' and A.activo=1 order by IsNull(A.fechamodifica,A.fecharegistra) desc,A.codemailnuevo desc"
						''Response.Write sql
						consultar sql,RS
						Do While Not RS.EOF						
						%>
						<tr style="display: none">
							<td bgcolor="#E9F8FE"><font size=2 face=Arial color=#00529B>&nbsp;<%=RS.Fields("email")%></font></td>
						</tr>							
						<%
						RS.MoveNext
						Loop
						RS.Close
						%>						
						</table>
				 </td>
			</tr>				
			</table>
			
			<%
            'CARGO CUOTAS
			sql="select A.CONTRATO,A.CODCENTRAL,A.CLIENTE,A.PRODUCTO,A.DA,A.FLAG_REF,A.CLASIFICACION,A.JUD,A.PROVCONST, " &_
		    "A.GARANT_TOTAL,A.TERRITORIO,A.OFICINA,A.FECHAPASEAMORA,A.SEGMENTO_RIESGO,A.Plaza,A.nomb_estu,A.ESPECIALISTA,A.FASIGNA,A.NUMPROC,A.DESC_PROCESO, " &_
		    "A.FECFASE,A.DESCRIPCION,B.DIAS_PER from FB.CentroCobranzas.dbo.PD_Detalle_Casos_Jud A inner join FB.CentroCobranzas.dbo.FASES_SAE B on A.codigo=B.CODIGO and A.TIPOPROCESO=B.TIPOPROCESO where A.CONTRATO='" & contrato & "' and A.FECHADATOS='" & fechacierre & "' and A.CODIGO is not null order by A.FECFASE desc" 
		    ''Response.Write sql
		    consultar sql,RS1
		    %>			
			<table width=100% cellpadding=0 cellspacing=0 border=0>		
			<tr bgcolor="#007DC5">
				<td height="22"><font size=2 face=Arial color="#FFFFFF"><b>&nbsp;Situación Judicial: (<%=RS1.RecordCount%>&nbsp;<%if RS1.RecordCount>1 then%>fases<%else%>fase<%end if%>)</b></font></td>
			</tr>
			</table>	
				

			<script language=javascript>
			    function visualizarfases() {
			        var filas = document.getElementById('tablafases').rows.length;
			        //if (document.getElementById('imagendir').src.length==document.getElementById('imagendir').src.replace('mostrar','').length)
			        if (document.getElementById('imagenfases').title == "Mostrar") {
			            document.getElementById('imagenfases').title = "Ocultar";
			            document.getElementById('imagenfases').alt = "Ocultar";
			            document.getElementById('imagenfases').src = "imagenes/ocultar.png";
			            for (i = 2; i < filas; i++) {
			                document.getElementById('tablafases').rows[i].style.display = '';
			            }
			        }
			        else {
			            document.getElementById('imagenfases').title = "Mostrar";
			            document.getElementById('imagenfases').alt = "Mostrar";
			            document.getElementById('imagenfases').src = "imagenes/mostrar.png";
			            for (i = 2; i < filas; i++) {
			                document.getElementById('tablafases').rows[i].style.display = 'none';
			            }
			        }
			    }
			</script>
				
			<table width=100% id="tablafases" cellpadding=1 cellspacing=1 border=0>
			<tr bgcolor="#007DC5">
			        <td width=25><font size=2 face=Arial color=#00529B>&nbsp;<a href="javascript:visualizarfases();"><img id="imagenfases" src="imagenes/mostrar.png" border=0 alt="Mostrar" title="Mostrar"></a></font></td>
			        <td width=10% align="left"><font size=2 face=Arial color="#FFFFFF"><b>N° Proceso</b></font></td>
					<td width=20% align="left"><font size=2 face=Arial color="#FFFFFF"><b>Proceso</b></font></td>
					<td width=10% align="left"><font size=2 face=Arial color="#FFFFFF"><b>Fecha</b></font></td>
					<td width=30% align="left"><font size=2 face=Arial color="#FFFFFF"><b>Fase</b></font></td>
					<td align="left"><font size=2 face=Arial color="#FFFFFF"><b>Días Permitidos</b></font></td>	
			</tr>
			<%
            nrofase=0			
			Do While not RS1.EOF
            nrofase=nrofase + 1
				'',(select count(distinct fechavencimiento) from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos) as NumCuotas,(select top 1 divisa from CuotaDiario where contrato=A.contrato and fechadatos=A.fechadatos and divisa<>A.divisa) as DivisaDif 
			%>
			<tr bgcolor="#E9F8FE" <%if nrofase<>1 then%> style="display: none"<%end if%>>	
			        <td><font size=2 face=Arial color=#00529B>&nbsp;</font></td>
					<td valign="top" align="left"><font size=2 face=Arial color=#00529B><%=RS1.Fields("NUMPROC")%></font></td>
					<td valign="top" align="left"><font size=2 face=Arial color=#00529B><%=RS1.Fields("DESC_PROCESO")%></font></td>
					<td valign="top" align="left"><font size=2 face=Arial color=#00529B><%=RS1.Fields("FECFASE")%></font></td>
					<td valign="top" align="left"><font size=2 face=Arial color=#00529B><%=RS1.Fields("DESCRIPCION")%></font></td>
					<td valign="top" align="left"><font size=2 face=Arial color=#00529B><%=RS1.Fields("DIAS_PER")%></font></td>
			</tr>
				
			<%
			RS1.MoveNext
			Loop
			RS1.Close
			%>
			</table>

			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Gestiones (<%=contadortotal%>)&nbsp;&nbsp;<!--<a href="javascript:agregargestion();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;--><a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
                <td bgcolor="#F5F5F5" align="right"><font size=2 face=Arial color=#00529B><b>Fecha inicio:</td>
			    <td bgcolor="#F5F5F5"><input name="fechagestionini" type=text maxlength=10 id="sel1" value="<%if IsDate(fechagestionini) then%><%=fechagestionini%><%else%><%=minfechagestion%><%end if%>" style="font-size: x-small; width: 70px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel1', '%d/%m/%Y');">
			    <td bgcolor="#F5F5F5" align="right"><font size=2 face=Arial color=#00529B><b>Fecha fin:</td>
			    <td bgcolor="#F5F5F5"><input name="fechagestionfin" type=text maxlength=10 id="sel2" value="<%if IsDate(fechagestionfin) then%><%=fechagestionfin%><%else%><%=maxfechagestion%><%end if%>" style="font-size: x-small; width: 70px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel2', '%d/%m/%Y');">
			    <td bgcolor="#F5F5F5"><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>		
			    <%if contador > 0 then%>		
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
			    </tr>
			    </table>
			    <div id="tabla0"> 
			    </div>
			    <%else%>
			    </tr>
			    </table>
			    <%end if%>
			
		<%''end if%>
			<input type="hidden" name="vistapadre" value="<%=obtener("vistapadre")%>">
			<input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
			<input type="hidden" name="contrato" value="<%=contrato%>">
		    <input type="hidden" name="expimp" value="">
		    <input type="hidden" name="pag" value="<%=pag%>">	
    	</form>
			
		
		<!--<script type="text/javascript">
		    initTriStateCheckBox('tristateBox1', 'tristateBox1State', true);
		</script>-->
		<%if contador > 0 then%>
		<script language="javascript">
		    inicio();
		</script>
		<%end if%>
		<!--cargando--><script language=javascript>		                   document.getElementById("imgloading").style.display = "none";</script>							
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
					consulta_exp="select 'Fecha de Gestión','Contrato','Agencia','Gestion','Comentario','F.Promesa','Direccion/Telefono','Adjunto'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select convert(varchar,A.fhgestionado,103) + ' ' + convert(varchar,A.fhgestionado,108),char(39) + A.contrato,C.RazonSocial,B.Descripcion, replace(replace(A.comentario,char(10),''),char(13),'') as comentario,ISNULL(convert(varchar,A.fechapromesa,103),''),CASE WHEN ltrim(rtrim(A.fono))<>'' THEN rtrim(ltrim(A.tipofono)) + '-' + rtrim(ltrim(A.prefijo)) + '-' + rtrim(ltrim(A.fono)) + (CASE WHEN LEN(RTRIM(ltrim(A.extension)))> 0 and A.extension<>'0000' THEN '-'+ RTRIM(LTRIM(A.extension)) ELSE '' END) ELSE rtrim(ltrim(A.direccion)) + ' - ' + rtrim(ltrim(UBI.distrito)) + ' - ' + rtrim(ltrim(UBI.provincia)) + ' - ' + rtrim(ltrim(UBI.departamento)) END,CASE WHEN A.ficherogestion IS NOT NULL THEN 'Sí' ELSE 'No' END " & _
								 "from CobranzaCM.dbo.RespuestaGestion A inner join CobranzaCM.dbo.Gestion B on A.codgestion=B.codgestion left outer join CobranzaCM.dbo.Agencia C on A.codagencia=C.codagencia left outer join CobranzaCM.dbo.Ubigeo UBI on UBI.coddpto=A.coddpto and UBI.codprov=A.codprov and UBI.coddist=A.coddist " & filtrobuscador & " order by A.fhgestionado desc"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					''response.Write sql
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
	    window.open("userexpira.asp", "_top");
	</script>
	<%	
	end if
	desconectar
else
%>
<script language="javascript">
    alert("Tiempo Expirado");
    window.open("index.html", "sistema");
    window.close();
</script>
<%
end if
%>



