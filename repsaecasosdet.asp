<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
	conectar

	if permisofacultad("repsaeestudio.asp") or permisofacultad("repsaeterritorio.asp") or permisofacultad("repsaegestor.asp")  or permisofacultad("repsaecasosdet.asp") or permisofacultad("repida.asp") or permisofacultad("repidasegmento.asp") or permisofacultad("repidasegmentopart.asp") or permisofacultad("repidaestudio.asp") then
	
	    paginapadre=obtener("paginapadre")
	    asignado=obtener("asignado")
	    codgestor=obtener("codgestor")
	    codestudio=obtener("codestudio")
	    codterritorio=obtener("codterritorio")
	    plaza=obtener("plaza")
	    segmento=obtener("segmento")
	    confase=obtener("confase")
	    contrato=obtener("contrato")
	    codigocentral=obtener("codigocentral")
	    tipoproceso=obtener("tipoproceso")
	    
	    fechacierre=obtener("fechacierre")
	    if fechacierre="" then
            sql="select max(fechadatos) as fechadatos from FB.CentroCobranzas.dbo.pd_casos_jud"
            consultar sql,RS	
            fechacierre=RS.Fields("fechadatos")
            RS.Close
		end if	    
	    
	    expimp=obtener("expimp")
	    flag=""
		
	    
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
	    
        filtrobuscador = ""
        filtrobuscador = " where A.segmento_riesgo in ('PARTICULARES','PYME') "
        
        if codterritorio<>"" then
            if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.codterritorio='" & codterritorio & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.codterritorio='" & codterritorio & "'"
			end if
		end if
        if codigocentral<>"" then
            if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.codcentral='" & codigocentral & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.codcentral='" & codigocentral & "'"
			end if
		end if
			
		if contrato<>"" then
            if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.contrato='" & contrato & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.contrato='" & contrato & "'"
			end if
		end if
        
		if codgestor<>"" then
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where replace(A.especialista,' ','')='" & codgestor & "'"
			else
			    filtrobuscador = filtrobuscador & " and replace(A.especialista,' ','')='" & codgestor & "'"
			end if
		end if      
		  
		  
		 ''TriState:
		 ''0 Sin asignar homologo asignado="2"
		 ''1 Todos       homologo asignado="0"
		 ''2 Asignados   homologo asignado="1"
		 
        select case asignado
            case "" : asignado="1"
            case "S" : asignado="2"
        end select

        select case asignado
	    case "0" : ''No Asignados
	                checkbuscasignados="value='0'" ''No Asignados
	                if filtrobuscador="" then
		                ''filtrobuscador = filtrobuscador & " where A.FASIGNA<=DATEADD(m,-2,getdate()) "
		                filtrobuscador = filtrobuscador & " where A.FASIGNA is null "
		            else
		                ''filtrobuscador = filtrobuscador & " and A.FASIGNA<=DATEADD(m,-2,getdate()) "
		                filtrobuscador = filtrobuscador & " and A.FASIGNA is null "
		            end if                    
	    case "2" : ''Asignados
				    checkbuscasignados="value='2'" ''Asignados
	                if filtrobuscador="" then
		                ''filtrobuscador = filtrobuscador & " where A.FASIGNA<=DATEADD(m,-2,getdate()) "
		                filtrobuscador = filtrobuscador & " where A.FASIGNA is not null "
		            else
		                ''filtrobuscador = filtrobuscador & " and A.FASIGNA<=DATEADD(m,-2,getdate()) "
		                filtrobuscador = filtrobuscador & " and A.FASIGNA is not null "
		            end if					    
	    case else: ''Todos
				    checkbuscasignados="value='1'"
	    end select	
			
		if codestudio<>"" then
		    flag="2"
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.CODDATO='" & codestudio & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.CODDATO='" & codestudio & "'"
			end if
		end if
		if plaza<>"" then
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.Plaza='" & plaza & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.Plaza='" & plaza & "'"
			end if
		end if		
		if segmento<>"" then
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.segmento_riesgo='" & segmento & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.segmento_riesgo='" & segmento & "'"
			end if
		end if			
		if confase<>"" then
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.codfase is not null"
			else
			    filtrobuscador = filtrobuscador & " and A.codfase is not null"
			end if
		end if	
		
		if fechacierre<>"" then
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.fechadatos='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.fechadatos='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'"
			end if
		end if		
		if tipoproceso<>"" then
		    if filtrobuscador="" then
			    filtrobuscador = filtrobuscador & " where A.tipoproceso='" & tipoproceso & "'"
			else
			    filtrobuscador = filtrobuscador & " and A.tipoproceso='" & tipoproceso & "'"
			end if
		end if			
			
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if				        
%>
<html>
<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
<head>
<title>SAE - Resumen de Casos</title>
        <script language=javascript src="scripts/TablaDinamica.js"></script>
		<script type="text/javascript" src="scripts/tristate-0.9.2.js" ></script>
		<script language=javascript>
		    var ventanaverimpagado;
		    function inicio() {
		        dibujarTabla(0);
		    }
		    function modificar(contrato) {
		        ventanavercasojudicial = window.open("vercasojudicial.asp?vistapadre=" + window.name + "&paginapadre=repsaecasosdet.asp&contrato=" + contrato + "&fechacierre=<%=fechacierre%>", "VerCasoJudicial" + contrato, "scrollbars=yes,scrolling=yes,top=" + ((screen.height) / 2 - 300) + ",height=600,width=" + (screen.width / 2 + 300) + ",left=" + (screen.width / 2 - 475) + ",resizable=yes");
		        ventanavercasojudicial.focus();
		    }
		    function actualizar() {
		        document.formula.actualizarlista.value = 1;
		        document.formula.submit();
		    }
		    function exportar() {
		        if (document.formula.buscando.value == "") {
		            document.formula.expimp.value = 1;
		            document.formula.submit();
		        }
		    }
		    function imprimir() {
		        window.open("impusuarios.asp", "ImpUsuarios", "scrollbars=yes,scrolling=yes,top=0,height=200,width=200,left=0,resizable=yes");
		    }
		    function buscar() {
		        if (document.formula.buscando.value == "") {
		            document.formula.buscando.value = "OK";
		            document.formula.pag.value = 1;
		            document.formula.submit();
		        }
		    }
		    function regresarpagpadre() 
		    {
		            document.formula.action = document.formula.paginapadre.value;
		            document.formula.submit();
		    }		    
		    function filtrar() {
		        if (filtrardatos[0] == 1) {
		            filtrardatos[0] = 0;
		            dibujarTabla(0);
		        }
		        else {
		            filtrardatos[0] = 1;
		            dibujarTabla(0);
		        }
		    }
		    function mostrarpag(pagina) {
		        if (document.formula.buscando.value == "") {
		            document.formula.buscando.value = "OK";
		            document.formula.pag.value = pagina;
		            document.formula.submit();
		        }
		    }
		</script>
		<style>
		A {
			FONT-SIZE: 10px; COLOR: #00529B; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
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
			font-size:10px;
			font-family:Arial;
			cursor:hand;
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
<script language=javascript>
    function actualizaproceso() {
        if (document.formula.buscando.value == "") {
            document.formula.actualizarlista.value = "";
            document.formula.pag.value = 1;
            document.formula.submit();
        }
    }
</script>

</head>
<%Response.Flush()%>
	<script language=javascript>
			rutaimgcab="imagenes/"; 
		  //Configuración general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=21;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '2';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('N° Contrato','Código','Nombre','Producto','D','Ref','Clasif','Deuda','Prov.Const.','Garantía Total','Territorio','Oficina','Pase a Mora','Segmento','Plaza','Estudio','Especialista','F.Asigna','Proceso','F.Fase','Fase');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array( true,true,true,true,true,true,false,true,false,false,true,false,false,true,true,true,true,true,true,true,true );
		    anchocolumna[tabla] =  new Array( '1%','1%','3%','4%','1%','1%','1%','1%','1%','1%','3%','1%','1%','1%','1%','3%','3%','1%','3%','1%','4%' );
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','middle','middle','left','right','right','right','left','left','left','left','left','left','left','left','left','left','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1,-1,-1,-1,0,-1,-1,2,2,2,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1,-1);
		    formatofecha[tabla] =   new Array('','','','','','','','','','','','','dd/mm/aaaa','','','','','dd/mm/aaaa','','dd/mm/aaaa','');

		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=contrato-id- value=-c0-><a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][1]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-c0-");>-valor-</a>';										
				objetofomulario[tabla][4]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][6]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][7]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][8]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][9]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][10]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][11]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][12]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][13]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][14]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][15]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][16]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][17]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][18]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][19]='<a href=javascript:modificar("-c0-");>-valor-</a>';
				objetofomulario[tabla][20]='<a href=javascript:modificar("-c0-");>-valor-</a>';
					
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
				filtrofomulario[tabla][12]='';
				filtrofomulario[tabla][13]='';
				filtrofomulario[tabla][14]='';
				filtrofomulario[tabla][15]='';
				filtrofomulario[tabla][16]='';
				filtrofomulario[tabla][17]='';
				filtrofomulario[tabla][18]='';
				filtrofomulario[tabla][19]='';	
				filtrofomulario[tabla][20]='';																
				
									
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
				valorfiltrofomulario[tabla][16]='';
				valorfiltrofomulario[tabla][17]='';
				valorfiltrofomulario[tabla][18]='';
				valorfiltrofomulario[tabla][19]='';	
				valorfiltrofomulario[tabla][20]='';													
				
			
		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		''response.write filtrobuscador		
		
		contadortotal=0
		sql="select count(*) from FB.CentroCobranzas.dbo.PD_Casos_Jud A " & filtrobuscador 
		consultar sql,RS	
		''response.write sql
		contadortotal=rs.fields(0)
		RS.Close
		
		cantidadxpagina=18
		paginasxbloque=10
		
		
		if obtener("pag")="" then
		pag=1
		else
		pag=int(obtener("pag"))
		end if


		if trim(contrato)<>"" then
		    xcontrato="'" & trim(contrato) & "'"
		else
		    xcontrato="NULL"
		end if
		if trim(codigocentral)<>"" then
		    xcodigocentral="'" & trim(codigocentral) & "'"
		else
		    xcodigocentral="NULL"
		end if		
		if codgestor<>"" then
		    xcodgestor=codgestor
		else
		    xcodgestor="NULL"
		end if	
		if codtipodocumento<>"" then
		    xcodtipodocumento="'" & codtipodocumento & "'"
		else
		    xcodtipodocumento="NULL"
		end if	
		if trim(numdocumento)<>"" then
		    xnumdocumento="'" & trim(numdocumento) & "'"
		else
		    xnumdocumento="NULL"
		end if		
		if isNumeric(diaatrasoini) then
		    xdiaatrasoini=int(diaatrasoini)
		else
		    xdiaatrasoini="NULL"
		end if	
		if isNumeric(diaatrasofin) then
		    xdiaatrasofin=int(diaatrasofin)
		else
		    xdiaatrasofin="NULL"
		end if	
        if isNumeric(codfamproducto) then
		    xcodfamproducto=int(codfamproducto)
		else
		    xcodfamproducto="NULL"
		end if			
		if isNumeric(codfamproducto) then
		    xcodfamproducto=int(codfamproducto)
		else
		    xcodfamproducto="NULL"
		end if		
		if isNumeric(codagencia) then
		    xcodagencia=int(codagencia)
		else
		    xcodagencia="NULL"
		end if			
		if isNumeric(codagencia) then
		    xcodagencia=int(codagencia)
		else
		    xcodagencia="NULL"
		end if	
		if isNumeric(codmarca) then
		    xcodmarca=int(codmarca)
		else
		    xcodmarca="NULL"
		end if				
        if codterritorio<>"" then
		    xcodterritorio="'" & codterritorio & "'"
		else
		    xcodterritorio="NULL"
		end if			
        if codoficina<>"" then
		    xcodoficina="'" & codoficina & "'"
		else
		    xcodoficina="NULL"
		end if					
        if not IsDate(fechapromesaini) then
		    xfechapromesaini="NULL"
		else
		    xfechapromesaini="'" & mid(fechapromesaini,7,4) & mid(fechapromesaini,4,2) & mid(fechapromesaini,1,2) & "'"
		end if
		if not IsDate(fechapromesafin) then
		    xfechapromesafin="NULL"
		else
		    xfechapromesafin="'" & mid(fechapromesafin,7,4) & mid(fechapromesafin,4,2) & mid(fechapromesafin,1,2) & "'"
		end if
        if codgestion<>"" then
		    xcodgestion="'" & codgestion & "'"
		else
		    xcodgestion="NULL"
		end if		
		
        ''@f_vistabusqueda as varchar(100),
        ''@f_cantidadxpagina as int,
        ''@f_pagina as int,
        ''@f_fechaasigini as datetime,
        ''@f_fechaasigfin as datetime,
        ''@f_contrato as varchar(18),
        ''@f_codigocentral as varchar(8),
        ''@f_codtipodocumento as varchar(1),
        ''@f_numdocumento as varchar(15),
        ''@f_diaatrasoini as int,
        ''@f_diaatrasofin as int,
        ''@f_codfamproducto as int,
        ''@f_codagencia as int,
        ''@f_codmarca as int,
        ''@f_codterritorio as varchar(4),
        ''@f_codoficina as varchar(4),
        ''@f_fechapromesaini as datetime,
        ''@f_fechapromesafin as datetime,
        ''@f_codgestion as varchar(3),
        ''@f_codgestor as int	
        	
	
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
		
		
		'''nuevo esquema
		if pag>1 then					
		sql="select top " & cantidadxpagina & " A.CONTRATO,A.CODCENTRAL,A.CLIENTE,A.PRODUCTO,A.DA,A.FLAG_REF,A.CLASIFICACION,A.JUD,A.PROVCONST,A.GARANT_TOTAL,A.TERRITORIO,A.OFICINA,A.FECHAPASEAMORA,A.SEGMENTO_RIESGO,A.Plaza,A.ESTUDIO,A.ESPECIALISTA,A.FASIGNA,A.PROCESO,A.FECFASE,IsNull(A.FASE,'') as FASE from FB.CentroCobranzas.dbo.PD_Casos_Jud A where " & filtrobuscador1 & " A.CONTRATO not in (select top " & topnovisible & " A.CONTRATO from FB.CENTROCOBRANZAS.dbo.PD_CASOS_JUD A order by A.CONTRATO) order by A.CONTRATO" 
		else
		sql="select top " & cantidadxpagina & " A.CONTRATO,A.CODCENTRAL,A.CLIENTE,A.PRODUCTO,A.DA,A.FLAG_REF,A.CLASIFICACION,A.JUD,A.PROVCONST,A.GARANT_TOTAL,A.TERRITORIO,A.OFICINA,A.FECHAPASEAMORA,A.SEGMENTO_RIESGO,A.Plaza,A.ESTUDIO,A.ESPECIALISTA,A.FASIGNA,A.PROCESO,A.FECFASE,IsNull(A.FASE,'') as FASE from FB.CentroCobranzas.dbo.PD_Casos_Jud A " & filtrobuscador & " order by A.CONTRATO" 
		end if		
		''response.write sql
		consultar sql,RS
		
		
		contador=0
		Do while not RS.EOF
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]='<%=RS.Fields("CONTRATO")%>';
			    datos[tabla][<%=contador%>][1]='<%=rs.Fields("CODCENTRAL")%>';
				datos[tabla][<%=contador%>][2]='<%if len(trim(replace(RS.Fields("CLIENTE"),"'","´")))<=15 then%><%=trim(replace(RS.Fields("CLIENTE"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("CLIENTE"),"'","´")),1,15) & "..."%><%end if%>';
				datos[tabla][<%=contador%>][3]='<%=rs.Fields("PRODUCTO")%>';
				datos[tabla][<%=contador%>][4]=<%=RS.Fields("DA")%>;
				datos[tabla][<%=contador%>][5]='<%=iif(RS.Fields("FLAG_REF")=0,"NO","SI")%>';
				datos[tabla][<%=contador%>][6]='<%=RS.Fields("CLASIFICACION")%>';
				datos[tabla][<%=contador%>][7]=<%=RS.Fields("JUD")%>;
				datos[tabla][<%=contador%>][8]=<%=RS.Fields("PROVCONST")%>;
				datos[tabla][<%=contador%>][9]=<%=RS.Fields("GARANT_TOTAL")%>;
				datos[tabla][<%=contador%>][10]='<%=RS.Fields("TERRITORIO")%>';
				datos[tabla][<%=contador%>][11]='<%=RS.Fields("OFICINA")%>';
				datos[tabla][<%=contador%>][12]=<%if not IsNull(RS.Fields("FECHAPASEAMORA")) then%>new Date(<%=Year(RS.Fields("FECHAPASEAMORA"))%>,<%=Month(RS.Fields("FECHAPASEAMORA"))-1%>,<%=Day(RS.Fields("FECHAPASEAMORA"))%>,<%=Hour(RS.Fields("FECHAPASEAMORA"))%>,<%=Minute(RS.Fields("FECHAPASEAMORA"))%>,<%=Second(RS.Fields("FECHAPASEAMORA"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][13]='<%=RS.Fields("SEGMENTO_RIESGO")%>';
				datos[tabla][<%=contador%>][14]='<%=RS.Fields("PLAZA")%>';
				datos[tabla][<%=contador%>][15]='<%if len(trim(replace(RS.Fields("ESTUDIO"),"'","´")))<=15 then%><%=trim(replace(RS.Fields("ESTUDIO"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("ESTUDIO"),"'","´")),1,15) & "..."%><%end if%>';
				datos[tabla][<%=contador%>][16]='<%if len(trim(replace(RS.Fields("ESPECIALISTA"),"'","´")))<=15 then%><%=trim(replace(RS.Fields("ESPECIALISTA"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("ESPECIALISTA"),"'","´")),1,15) & "..."%><%end if%>';			
				datos[tabla][<%=contador%>][17]=<%if not IsNull(RS.Fields("FASIGNA")) then%>new Date(<%=Year(RS.Fields("FASIGNA"))%>,<%=Month(RS.Fields("FASIGNA"))-1%>,<%=Day(RS.Fields("FASIGNA"))%>,<%=Hour(RS.Fields("FASIGNA"))%>,<%=Minute(RS.Fields("FASIGNA"))%>,<%=Second(RS.Fields("FASIGNA"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][18]='<%=RS.Fields("PROCESO")%>';
				datos[tabla][<%=contador%>][19]=<%if not IsNull(RS.Fields("FECFASE")) then%>new Date(<%=Year(RS.Fields("FECFASE"))%>,<%=Month(RS.Fields("FECFASE"))-1%>,<%=Day(RS.Fields("FECFASE"))%>,<%=Hour(RS.Fields("FECFASE"))%>,<%=Minute(RS.Fields("FECFASE"))%>,<%=Second(RS.Fields("FECFASE"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][20]='<%if len(trim(replace(RS.Fields("FASE"),"'","´")))<=15 then%><%=trim(replace(RS.Fields("FASE"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("FASE"),"'","´")),1,15) & "..."%><%end if%>';			
		<%
			contador=contador + 1
			RS.MoveNext 
		Loop 
		RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','','','','','','','','','','','','','','','','',''); 


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
        objetosdebusqueda="<span id='tristateBox1' style='cursor: default;'><input type='hidden' id='tristateBox1State' name='asignado' " & checkbuscasignados & "></span>"
        %>	

<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
<form name=formula method=post action="repsaecasosdet.asp">
<table border=0 cellspacing=0 cellpadding=0 width=100%>
<tr>	
	<td bgcolor="#F5F5F5" align="right" width=40%>			
		<font size=4 color=#483d8b face=Arial><%if paginapadre<>"" then%><a href="javascript:regresarpagpadre();"><img src="imagenes/atras.png" border=0 alt="Buscar" title="Buscar" align=middle height="18"></a><%end if%></font>
	</td>
	<td bgcolor="#F5F5F5" align="left">			
		<font size=4 color=#483d8b face=Arial>&nbsp;&nbsp;<b>SAE - Resumen de Casos al <%=fechacierre%></b></font>
	</td>
</tr>
</table>
<table border=0 cellspacing=0 cellpadding=2 width=100%>
<tr>
  <td>
		 <%if contador=0 then%>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=1 face=Arial color=#00529B><b>Casos (0) - No hay registros.</b></font>&nbsp;</td>
			</tr>
			</table>
		<%else		
		%>
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=1 face=Arial color=#00529B><b>Casos (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B><b>Asignados:</b></font></td>
	            <td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
	            <td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B><b>N°&nbsp;Proceso:</b></font></td>
	            <td bgcolor="#F5F5F5">
	                <select name="TIPOPROCESO" style="font-size: x-small; width: 250px;">
			        <option value="">Seleccione Proceso</option>
			        <%
				        sql = "select distinct A.TIPOPROCESO,A.DESC_PROCESO FROM FB.CentroCobranzas.dbo.FASES_SAE A ORDER BY TIPOPROCESO"
				        consultar sql,RS
				        Do While Not  RS.EOF
				        %>
				        <option value="<%=RS.Fields("TIPOPROCESO")%>" <%if TIPOPROCESO<>"" then%><% if RS.fields("TIPOPROCESO")=TIPOPROCESO then%> selected<%end if%><%end if%>><%=RS.Fields("DESC_PROCESO")%></option>
				        <%
				        RS.MoveNext
				        loop
				        RS.Close
			        %>
		            </select>
	            </td>
	            <td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B><b>N°&nbsp;Contrato:</b></font></td>
                <td bgcolor="#F5F5F5" align=right><input name="contrato" type=text maxlength=18 value="<%=contrato%>" style="font-size: x-small; width: 130px;"></td>
                <td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
                <td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B><b>Código&nbsp;Cliente:</b></font></td>
                <td bgcolor="#F5F5F5" align=right><input name="codigocentral" type=text maxlength=8 value="<%=codigocentral%>" style="font-size: x-small; width: 70px;"></td>
                <td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
                 <td bgcolor="#F5F5F5" align=right><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=right></a></td>		
				<td bgcolor="#F5F5F5" align=right><font size=1 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
			</tr>	
			</table>
			<div id="tabla0"> 
			</div>
		<%end if%>
        <input type="hidden" name="buscando" value="">				
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="pag" value="<%=pag%>">
	    <input type="hidden" name="paginapadre" value="<%=obtener("paginapadre")%>">
	    <input type="hidden" name="codgestor" value="<%=obtener("codgestor")%>">
   	    <input type="hidden" name="codestudio" value="<%=obtener("codestudio")%>">
        <input type="hidden" name="plaza" value="<%=obtener("plaza")%>">
	    <input type="hidden" name="segmento" value="<%=obtener("segmento")%>">
	    <input type="hidden" name="confase" value="<%=obtener("confase")%>">	    	    	    	    
	    <input type="hidden" name="fechacierre" value="<%=fechacierre%>">    	    	    	    
					
		<script type="text/javascript">
		    initTriStateCheckBox('tristateBox1', 'tristateBox1State', true);
        </script>
		<script language="javascript">
		    inicio();
		</script>					
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display = "none";</script>		
		<%
		''Codigo exp excel
		''Si se pide exportar a excel
				if expimp="1" then
				
                                            
                        
		            ''response.write "<!--" & sql & "-->"
		            ''conn.execute sql                 
          		 
                    		                     
                    ''end if
                   '' RS.Close

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
					consulta_exp="select 'Contrato', 'Cod.Central','Nombre','Producto','Días de Atraso','Refinanciado','Clasificación','Deuda Judicial','Provisión','Garantía','Territorio','Oficina','Fecha de pase a mora','Segmento','Plaza','Estudio','Especialista','Fecha de Asignación','Proceso','Fecha de fase','Fase'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt

                    consulta_exp="select char(39) + A.CONTRATO,char(39) + A.CODCENTRAL,A.CLIENTE,A.PRODUCTO,A.DA,A.FLAG_REF,A.CLASIFICACION,A.JUD,A.PROVCONST,A.GARANT_TOTAL,A.TERRITORIO,A.OFICINA,convert(varchar,A.FECHAPASEAMORA,103)  AS FECHAPASEAMORA,A.SEGMENTO_RIESGO,A.Plaza,A.ESTUDIO,A.ESPECIALISTA,convert(varchar,A.FASIGNA,103) AS FASIGNA,A.PROCESO,convert(varchar,A.FECFASE,103)  AS FECFASE,A.FASE from FB.CentroCobranzas.dbo.PD_Casos_Jud A  " & filtrobuscador & " order by A.CONTRATO " 
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''consulta_exp="select char(39) + A.codigocentral,char(39) + A.contrato,A.nombre,char(39) + CASE WHEN A.tipofono1='C' THEN A.fono1 ELSE CASE WHEN rtrim(ltrim(A.prefijo1))='001' THEN CASE WHEN rtrim(ltrim(A.extension1))<>'0000' and rtrim(ltrim(A.extension1))<>'' THEN A.fono1 + '.Ext.' + A.extension1 ELSE A.fono1 END ELSE CASE WHEN rtrim(ltrim(A.extension1))<>'0000' and rtrim(ltrim(A.extension1))<>'' THEN A.prefijo1 + A.fono1 + '.Ext.' + A.extension1 ELSE A.prefijo1 + A.fono1 END END END as telefono,A.producto,A.subproducto,A.diasvencimiento,A.divisa,A.saldohoy,A.Agencia,A.fechagestion,A.fhgestionado,A.codgestion + ' - ' + A.Gestion as Gestion,A.TipoContacto " & _
					''			"from CobranzaCM.dbo." & vistabusqueda & " A " & filtrobuscador & " order by A.nro"								 
					''sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					''response.Write sql
					''conn.execute sql

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
  </td>
  </tr>   
</table>
</form>

</body>
<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display = "none";</script>
</html>
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
    window.open("index.html", "_top");
</script>
<%
end if
%>

