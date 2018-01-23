<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
	conectar
	''traemos maxima fecha de gestion en distribución diaria
	''si se seleccions fecha de asignacion inicio y fin igual a esta maxima fecha de gestion
	''se cruza con la tabla UltClienteDiario y UltContratoDiario para rapidez de reporte de gestion
	sql="select max(fechagestion) from DistribucionDiaria"
	consultar sql,RS	
	maxfechagestion=rs.fields(0)
	RS.Close
	
	sql="select A.codagencia,B.razonsocial as Agencia,A.codoficina,C.descripcion as Oficina,D.codterritorio,D.descripcion as Territorio from Usuario A left outer join Agencia B on A.codagencia=B.codagencia left outer join Oficina C on A.codoficina=C.codoficina left outer join Territorio D on C.codterritorio=D.Codterritorio where A.codusuario = " & session("codusuario")
	consultar sql,RS	
	codagenciausuario=rs.fields("codagencia")
	agenciausuario=rs.fields("agencia")
	codoficinausuario=rs.fields("codoficina")
	oficinausuario=rs.fields("oficina")
	codterritoriousuario=rs.fields("codterritorio")
	territoriousuario=rs.fields("territorio")	
	rs.close	
	
	actualizarlista=obtener("actualizarlista")
	
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
			
	if permisofacultad("admimpagado.asp") then
		contrato=obtener("contrato1")
		codigocentral=obtener("codigocentral1")
		codgestor=obtener("codgestor1")
		codtipodocumento=obtener("codtipodocumento1")
		numdocumento=obtener("numdocumento1")
		diaatrasoini=obtener("diaatrasoini1")
		diaatrasofin=obtener("diaatrasofin1")
		codfamproducto=obtener("codfamproducto1")
		codagencia=obtener("codagencia1")
		codmarca=obtener("codmarca1")
		if obtener("actualizarlista")<>"" then
            codterritorio=obtener("codterritorio1")
            if codterritorio="" then
	            codoficina=""
            else
	            codoficina=obtener("codoficina1")
            end if				
        else
            codterritorio=obtener("codterritorio")
            if codterritorio="" then
	            codoficina=""
            else
	            codoficina=obtener("codoficina")
            end if				        
        end if		
		tipogestion=obtener("tipogestion1")
		fechaasigini=obtener("fechaasigini1")
		fechaasigfin=obtener("fechaasigfin1")
		fechapromesaini=obtener("fechapromesaini1")
		fechapromesafin=obtener("fechapromesafin1")
		codgestion=obtener("codgestion1")
		if not IsDate(fechasigini) then
		    fechasigini=CStr(maxfechagestion)
		end if
		if not IsDate(fechasigfin) then
		    fechasigfin=CStr(maxfechagestion)
		end if		

        
%>
<html>
<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
<head>
<title>Seguimiento de Impagado</title>
<script language=javascript src="scripts/TablaDinamica.js"></script>
		<script language=javascript>
		var ventanaverimpagado;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codcen,contr,fd,fg)
		{
			ventanaverimpagado=window.open("verimpagado.asp?vistapadre=" + window.name + "&paginapadre=admimpagado.asp&codigocentral=" + codcen + "&contrato=" + contr + "&fechadatos=" + fd + "&fechagestion=" + fg,"VerImpagado" + codcen,"scrollbars=yes,scrolling=yes,top=" + ((screen.height)/2 - 300) + ",height=600,width=" + (screen.width/2 + 300) + ",left=" + (screen.width/2 - 475) + ",resizable=yes");
			ventanaverimpagado.focus();
		}			
		function actualizar()
		{
			document.formula.actualizarlista.value=1;
			document.formula.submit();
		}	
		function exportar()
		{
		    if (document.formula.buscando.value == "") 
		    {		
			document.formula.expimp.value=1;
			document.formula.submit();
			}
		}	
		function imprimir()
		{
			window.open("impusuarios.asp","ImpUsuarios","scrollbars=yes,scrolling=yes,top=0,height=200,width=200,left=0,resizable=yes");
		}					
		function buscar()
		{
		    if (document.formula.buscando.value == "") 
		    {
		        document.formula.buscando.value = "OK";
		        document.formula.nuevabusqueda.value = 1;
		        document.formula.actualizarlista.value = 1;
		        document.formula.pag.value = 1;
		        document.formula.contrato1.value = document.formula.contrato.value;
		        document.formula.codigocentral1.value = document.formula.codigocentral.value;
		        document.formula.codgestor1.value = document.formula.codgestor.value;
		        document.formula.codtipodocumento1.value = document.formula.codtipodocumento.value;
		        document.formula.numdocumento1.value = document.formula.numdocumento.value;
		        document.formula.diaatrasoini1.value = document.formula.diaatrasoini.value;
		        document.formula.diaatrasofin1.value = document.formula.diaatrasofin.value;
		        document.formula.codfamproducto1.value = document.formula.codfamproducto.value;
		        document.formula.codagencia1.value = document.formula.codagencia.value;
		        document.formula.codmarca1.value = document.formula.codmarca.value;
		        document.formula.codterritorio1.value = document.formula.codterritorio.value;
		        document.formula.codoficina1.value = document.formula.codoficina.value;
		        document.formula.tipogestion1.value = document.formula.tipogestion.value;
		        document.formula.fechaasigini1.value = document.formula.fechaasigini.value;
		        document.formula.fechaasigfin1.value = document.formula.fechaasigfin.value;
		        document.formula.fechapromesaini1.value = document.formula.fechapromesaini.value;
		        document.formula.fechapromesafin1.value = document.formula.fechapromesafin.value;
		        document.formula.codgestion1.value = document.formula.codgestion.value;
		        document.formula.submit();
		    }
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
			if (document.formula.buscando.value == "") 
		    {
		        document.formula.buscando.value = "OK";
			    document.formula.pag.value=pagina;
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
		

 
<script language=javascript>
	function actualizaterritorio()
	{
		    if (document.formula.buscando.value == "") 
		    {	
		        document.formula.actualizarlista.value="";
		        document.formula.pag.value=1;
		        document.formula.submit();
		    }
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
<%Response.Flush()%>
	<%if actualizarlista<>"" then
		if fechaasigini=fechaasigfin and fechaasigini=CStr(maxfechagestion) then
			vistabusqueda="VISTAULTIMPAGADO"
		else
			vistabusqueda="VISTAIMPAGADO"
		end if
	%>
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
			paddingtabla[tabla] = '2';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('Código','N° Contrato','Nombre','Teléfono','Producto','Subprod','D','Mon','Deuda','Agencia','F.Asig','F.Gest','Mejor&nbsp;Gestión','Tipo Contacto','Fecha Datos','Fecha Gestion');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array( true, true, true,true,true,true, true, true,true,true,true, true, true,true,false,false);
		    anchocolumna[tabla] =  new Array( '1%' , '1%' ,  '4%',  '1%',  '3%','3%' , '1%' ,  '1%',  '1%',  '2%','1%' , '1%' ,  '5%',  '3%',  '',  '');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','center','left','left','left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left','right','center','right','left','left','left','left','left','left','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1   ,-1 ,-1 ,-1,-1 ,0   ,-1 ,2 ,-1,-1 ,-1   ,-1 ,-1 ,-1,-1);
		    formatofecha[tabla] =   new Array(''  ,''  ,'' ,'','',''  ,''  ,'' ,'','','dd/mm/aaaa'  ,'dd/mm&nbsp;HH:MI'  ,'' ,'','','');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=codigocentral-id- value=-c0-><a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][1]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';											
				objetofomulario[tabla][4]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][6]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][7]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][8]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][9]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][10]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][11]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][12]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][13]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][14]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
				objetofomulario[tabla][15]='<a href=javascript:modificar("-c0-","-c1-","-c14-","-c15-");>-valor-</a>';
					
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
		<%
				
        filtrobuscador = " where DD.fechagestion between convert(datetime,'" & mid(fechaasigini,7,4) & mid(fechaasigini,4,2) & mid(fechaasigini,1,2) & "') and convert(datetime,'" & mid(fechaasigfin ,7,4) & mid(fechaasigfin,4,2) & mid(fechaasigfin,1,2) & "')"
		if contrato<>"" then
			filtrobuscador = filtrobuscador & " and DD.contrato='" & contrato & "'"
		end if
		if codigocentral<>"" then
			filtrobuscador = filtrobuscador & " and DD.codigocentral='" & codigocentral & "'"
		end if		
		if codtipodocumento<>"" then
			filtrobuscador = filtrobuscador & " and DD.tipodocumento='" & codtipodocumento & "'"
		end if		
		if numdocumento<>"" then
			filtrobuscador = filtrobuscador & " and DD.numdocumento='" & numdocumento & "'"
		end if		
		if diaatrasoini<>"" then
			filtrobuscador = filtrobuscador & " and DD.diasvencimiento>=" & diaatrasoini
		end if				
		if diaatrasofin<>"" then
			filtrobuscador = filtrobuscador & " and DD.diasvencimiento<=" & diaatrasofin
		end if						
		if codfamproducto<>"" then
			filtrobuscador = filtrobuscador & " and FP.codfamproducto=" & codfamproducto
		end if								
		if codagencia<>"" then
			filtrobuscador = filtrobuscador & " and DD.codagencia=" & codagencia
		end if										
		if codmarca<>"" then
			filtrobuscador = filtrobuscador & " and DD.codigocentral in (select codigocentral from MarcaCliente where codmarca=" & codmarca & " and activo=1)"
		end if												
		if codterritorio<>"" then
			filtrobuscador = filtrobuscador & " and DD.codterritorio='" & codterritorio & "'"
		end if												
		if codoficina<>"" then
			filtrobuscador = filtrobuscador & " and DD.codoficina='" & codoficina & "'"
		end if										
		''Preventiva/Impagado				
		''if tipogestion<>"" then
		''	filtrobuscador = filtrobuscador & " and A.tipogestion='" & tipogestion & "'"
		''end if																
		if fechapromesaini<>"" and fechapromesafin<>"" then
			filtrobuscador = " where RG.fechapromesa between '" & mid(fechapromesaini,7,4) & mid(fechapromesaini,4,2) & mid(fechapromesaini,1,2) & "' and '" & mid(fechapromesafin,7,4) & mid(fechapromesafin,4,2) & mid(fechapromesafin,1,2) & "'"
		end if												
		if fechapromesaini<>"" and fechapromesafin="" then
			filtrobuscador = " where RG.fechapromesa>='" & mid(fechapromesaini,7,4) & mid(fechapromesaini,4,2) & mid(fechapromesaini,1,2) & "' and '" & mid(fechapromesaini,7,4) & mid(fechapromesaini,4,2) & mid(fechapromesaini,1,2) & "'"
		end if														
		if fechapromesaini="" and fechapromesafin<>"" then
			filtrobuscador = " where RG.fechapromesa<='" & mid(fechapromesafin,7,4) & mid(fechapromesafin,4,2) & mid(fechapromesafin,1,2) & "'"
		end if																
		if codgestion<>"" then
			if codgestion<>"*" and codgestion<>"**" then
			filtrobuscador = filtrobuscador & " and RG.codgestion='" & codgestion & "'"
			end if
			if codgestion="*" then
			filtrobuscador = filtrobuscador & " and RG.codgestion is null"
			end if
			if codgestion="**" then
			filtrobuscador = filtrobuscador & " and RG.codgestion is not null"
			end if
		end if																	
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if		
		
		''response.write filtrobuscador		
		
		contadortotal=0
		
		sql="select OBJECT_ID('tempdb.dbo.##User" & session("codusuariodif") & "')"
	    consultar sql,RS
		 
        if obtener("nuevabusqueda")<>"" or IsNull(RS.Fields(0)) then
            if not IsNull(RS.Fields(0)) then
                sql="drop table tempdb.dbo.##User" & session("codusuariodif")
                conn.execute sql
            end if
            sql1="select DD.codigocentral, " &_
                    "DD.contrato," &_ 
                    "DD.fechadatos, " &_
                    "DD.codagencia," &_ 
                    "DD.nombre," &_
                    "DD.codterritorio," &_
                    "TE.descripcion AS Territorio," &_
                    "DD.codoficina," &_
                    "OFI.descripcion AS oficina," &_
                    "DD.telefono, " &_
                    "DD.tipofono1, " &_
                    "DD.prefijo1, " &_
                    "DD.fono1, " &_
                    "DD.extension1," &_
                    "CASE WHEN DD.direccion <>'' THEN DD.direccion + '-' + UBI2.departamento + '-' + UBI2.provincia + '-' + UBI2.distrito ELSE ' ' END as direccion," &_
                    "DD.CodProducto," &_
                    "DD.CodSubProducto," &_
                    "IsNull(SP.descripcion,'') as SubProducto," &_
                    "IsNull(FP.descripcion,'') as Producto," &_
                    "DD.Divisa," &_ 
                    "DD.SaldoHoy, " &_
                    "CASE WHEN DD.divisa <>'PEN' THEN DD.SaldoHoy*TCA.fixing ELSE DD.SaldoHoy END AS SaldoHoySoles," &_
                    "CASE WHEN DD.saldohoy*TCA.fixing <= 10 THEN 'RI1 [<10]' WHEN DD.saldohoy * TCA.fixing <= 100 THEN 'RI2 <10 - 100]' WHEN DD.saldohoy * TCA.fixing <= 200 THEN 'RI3 <100 - 200]' WHEN DD.saldohoy * TCA.fixing <= 300 THEN 'RI4 <200 - 300]' WHEN DD.saldohoy * TCA.fixing <= 500 THEN 'RI5 <300 - 500]' WHEN DD.saldohoy * TCA.fixing <= 1000 THEN 'RI6 <500 - 1000]' WHEN DD.saldohoy * TCA.fixing <= 1500 THEN 'RI7 <1000 - 1500]' WHEN DD.saldohoy * TCA.fixing <= 2000 THEN 'RI8 <1500 - 2000]' WHEN DD.saldohoy * TCA.fixing <= 2500 THEN 'RI9 <2000 - 2500]' WHEN DD.saldohoy * TCA.fixing <= 3000 THEN 'RI10 <2500 - 3000]' WHEN DD.saldohoy * TCA.fixing <= 3500 THEN 'RI11 <3000 - 3500]' WHEN DD.saldohoy * TCA.fixing <= 4000 THEN 'RI12 <3500 - 4000]' WHEN DD.saldohoy * TCA.fixing <= 5000 THEN 'RI13 <4000 - 5000]' WHEN DD.saldohoy * TCA.fixing <= 6000 THEN 'RI14 <5000 - 6000]' WHEN DD.saldohoy * TCA.fixing <= 7000 THEN 'RI15 <6000 - 7000]' WHEN DD.saldohoy * TCA.fixing <= 8000 THEN 'RI16 <7000 - 8000]' WHEN  DD.saldohoy * TCA.fixing <= 9000 THEN 'RI17 <8000 - 9000]' WHEN DD.saldohoy * TCA.fixing <= 10000 THEN 'RI18 <9000 - 10000]' WHEN DD.saldohoy * TCA.fixing <= 11000 THEN 'RI19 <10000 - 11000]' WHEN DD.saldohoy * TCA.fixing <= 12000 THEN 'RI20 <11000 - 12000]' WHEN DD.saldohoy * TCA.fixing <= 13000 THEN 'RI21 <12000 - 13000]'  WHEN DD.saldohoy * TCA.fixing > 13000 THEN 'RI22 [>13000]' END AS RangoImpago," &_
                    "DD.DiasVencimiento, " &_
                    "CASE WHEN DD.DiasVencimiento <= 30 THEN 'T1 [1-30]' WHEN DD.DiasVencimiento <= 60 THEN 'T2 [31-60]' WHEN DD.DiasVencimiento <= 90 THEN 'T3 [61-90]' WHEN DD.DiasVencimiento <= 120 THEN 'T4 [91-120]' END AS TramoReal," &_
                    "MA.descripcion AS Marca," &_
                    "AG.razonsocial as Agencia, " &_
                    "DD.fechaincumplimiento," &_
                    "DD.FechaGestion, " &_
                    "RG.fhgestionado, " &_
                    "GS.codgestion," &_
                    "GS.CodGestion + '-' + GS.descripcion as Gestion," &_
                    "IsNull(GS.codtipocontacto,'') as CodTipoContacto," &_
                    "IsNull(TC.descripcion,'') as TipoContacto," &_
                    "IsNull(convert(varchar,RG.FechaPromesa,103),'') + ' ' + IsNull(convert(varchar,RG.FechaPromesa,108),'') AS fechapromesa," &_
                    "CASE WHEN RG.fono<>'' THEN RG.tipofono + '-' + RG.prefijo + '-' + RG.fono + (CASE WHEN LTRIM(RTRIM(RG.extension))<>'' and LTRIM(RTRIM(RG.extension))<>'0000' THEN '-' + LTRIM(RTRIM(RG.Extension)) END )ELSE ' ' END AS telefonoGestion," &_
                    "CASE WHEN RG.direccion <>'' THEN RG.direccion + '-' + UBI.departamento + '-' + UBI.provincia + '-' + UBI.distrito ELSE ' ' END as direcciongestion," &_
                    "RG.Divisa1 AS DivisaPromesa1, " &_
                    "RG.Importe1," &_ 
                    "RG.Divisa2 AS DivisaPromesa2, " &_
                    "RG.Importe2," &_
                    "DD.tipodocumento, " &_
                    "DD.numdocumento, " &_
                    "DD.Nro" &_
                    " into ##User" & session("codusuariodif") & " from " &_
                    "BusquedaDistribucion DD inner join Agencia AG on DD.codagencia=AG.codagencia " &_
                    "left outer join Producto PR on DD.codproducto=PR.codproducto " &_
                    "left outer join SubProducto SP on DD.codsubproducto=SP.codsubproducto " &_
                    "left outer join FamProducto FP ON SP.codfamproducto=FP.codfamproducto " &_
                    "left outer join RespuestaGestion RG on DD.codigocentral=RG.codigocentral and DD.contrato=RG.contrato and DD.mejor_codrespgestion=RG.codrespgestion " &_
                    "left outer join Gestion GS ON RG.codgestion = GS.codgestion " &_
                    "left outer join TipoContacto TC ON GS.codtipocontacto = TC.codtipocontacto " &_
                    "left outer join Oficina OFI ON DD.codoficina=OFI.codoficina " &_
                    "left outer join Territorio TE ON DD.CodTerritorio=TE.codterritorio " &_
                    "left outer join TipoCambio TCA ON DD.fechadatos=TCA.fechadatos AND DD.Divisa=TCA.divisa and TCA.tipo='S' " &_
                    "left outer join MarcaCliente MC ON DD.codigocentral=MC.codigocentral and MC.activo=1 " &_
                    "left outer join Marca MA ON MC.codmarca=MA.codmarca " &_
                    "left outer join Ubigeo UBI on RG.coddpto=UBI.coddpto and RG.coddist=UBI.coddist and RG.codprov=UBI.codprov " &_
                    "left outer join Ubigeo UBI2 on DD.codestado=UBI2.coddpto and DD.coddistrito=UBI2.coddist and DD.codprovincia=UBI2.codprov " & filtrobuscador
		    conn.execute sql1            
        end if
        RS.Close
        
        ''sql1="select DD.codigocentral, DD.contrato,DD.fechadatos,DD.codagencia,DD.nombre,DD.telefono,DD.tipofono1,DD.prefijo1,DD.fono1,DD.extension1,DD.codproducto, DD.CodSubProducto,SP.descripcion as SubProducto, FP.descripcion as Producto, DD.DiasVencimiento,DD.Divisa,DD.SaldoHoy,AG.razonsocial as Agencia,DD.FechaGestion,RG.fhgestionado,RG.CodGestion,GS.descripcion as Gestion,GS.CodTipoContacto,TC.descripcion as TipoContacto,RG.FechaPromesa,RG.Divisa1 AS DivisaPromesa1,RG.Importe1,RG.Divisa2 AS DivisaPromesa2, RG.Importe2,DD.tipodocumento,DD.numdocumento,DD.CodTerritorio,DD.CodOficina, DD.Nro into ##User" & session("codusuariodif") & " from BusquedaDistribucion DD inner join Agencia AG on DD.codagencia=AG.codagencia left outer join Producto PR on DD.codproducto=PR.codproducto left outer join SubProducto SP on DD.codsubproducto=SP.codsubproducto left outer join FamProducto FP ON SP.codfamproducto=FP.codfamproducto left outer join RespuestaGestion RG on DD.codigocentral=RG.codigocentral and DD.contrato=RG.contrato and DD.mejor_codrespgestion=RG.codrespgestion left outer join Gestion GS ON RG.codgestion = GS.codgestion left outer join TipoContacto TC ON GS.codtipocontacto = TC.codtipocontacto left outer join Oficina OFI ON DD.codoficina=OFI.codoficina left outer join Territorio TE ON DD.CodTerritorio=TE.codterritorio left outer join TipoCambio TCA ON DD.fechadatos=TCA.fechadatos AND DD.Divisa=TCA.divisa and TCA.tipo='S' left outer join MarcaCliente MC ON DD.codigocentral=MC.codigocentral and MC.activo=1 left outer join Marca MA ON MC.codmarca=MA.codmarca left outer join Ubigeo UBI on RG.coddpto=UBI.coddpto and RG.coddist=UBI.coddist and RG.codprov=UBI.codprov " & filtrobuscador
        ''response.write sql1
		''conn.execute sql1
		
		sql="select count(*) from ##User" & session("codusuariodif") 
		''response.write sql
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
		sql="select top " & cantidadxpagina & " * from ##User" & session("codusuariodif") & " A where A.nro not in (select top " & topnovisible & " A.nro from ##User" & session("codusuariodif") & " A order by A.nro) order by A.nro" 
		else
		sql="select top " & cantidadxpagina & " * from ##User" & session("codusuariodif") & " A order by A.nro" 
		end if		
		''response.write sql
		consultar sql,RS
		
		contador=0
		Do while not RS.EOF
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]='<%=RS.Fields("codigocentral")%>';
			    datos[tabla][<%=contador%>][1]='<%=rs.Fields("contrato")%>';
				datos[tabla][<%=contador%>][2]='<%if len(trim(replace(RS.Fields("nombre"),"'","´")))<=27 then%><%=trim(replace(RS.Fields("nombre"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("nombre"),"'","´")),1,27) & "..."%><%end if%>';
				datos[tabla][<%=contador%>][3]='<%=iif(trim(RS.Fields("tipofono1"))="C",RS.Fields("fono1"),iif(trim(RS.Fields("prefijo1"))="001",iif(trim(RS.Fields("extension1"))<>"0000" and trim(RS.Fields("extension1"))<>"",RS.Fields("fono1") & ".Ext." & RS.Fields("extension1"),RS.Fields("fono1")),iif(trim(RS.Fields("extension1"))<>"0000" and trim(RS.Fields("extension1"))<>"",RS.Fields("prefijo1") & RS.Fields("fono1") & ".Ext." & RS.Fields("extension1"),RS.Fields("prefijo1") & RS.Fields("fono1"))))%>';
				datos[tabla][<%=contador%>][4]='<%=RS.Fields("producto")%>';
				datos[tabla][<%=contador%>][5]='<%if len(trim(replace(RS.Fields("subproducto"),"'","´")))<=20 then%><%=trim(replace(RS.Fields("subproducto"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("subproducto"),"'","´")),1,20) & "..."%><%end if%>';
				datos[tabla][<%=contador%>][6]=<%=RS.Fields("diasvencimiento")%>;
				datos[tabla][<%=contador%>][7]='<%=RS.Fields("divisa")%>';
				datos[tabla][<%=contador%>][8]=<%=RS.Fields("saldohoy")%>;
				datos[tabla][<%=contador%>][9]='<%if len(trim(RS.Fields("agencia")))<=15 then%><%=trim(RS.Fields("agencia"))%><%else%><%=mid(trim(RS.Fields("agencia")),1,15) & "..."%><%end if%>';
				datos[tabla][<%=contador%>][10]=<%if not IsNull(RS.Fields("fechagestion")) then%>new Date(<%=Year(RS.Fields("fechagestion"))%>,<%=Month(RS.Fields("fechagestion"))-1%>,<%=Day(RS.Fields("fechagestion"))%>,<%=Hour(RS.Fields("fechagestion"))%>,<%=Minute(RS.Fields("fechagestion"))%>,<%=Second(RS.Fields("fechagestion"))%>)<%else%>null<%end if%>;
				datos[tabla][<%=contador%>][11]=<%if not IsNull(RS.Fields("fhgestionado")) then%>new Date(<%=Year(RS.Fields("fhgestionado"))%>,<%=Month(RS.Fields("fhgestionado"))-1%>,<%=Day(RS.Fields("fhgestionado"))%>,<%=Hour(RS.Fields("fhgestionado"))%>,<%=Minute(RS.Fields("fhgestionado"))%>,<%=Second(RS.Fields("fhgestionado"))%>)<%else%>null<%end if%>;
				//datos[tabla][<%=contador%>][12]='<%=iif(RS.Fields("codgestion")<>"",iif(len(trim(RS.Fields("codgestion") & " - " & RS.Fields("gestion")))<=30,trim(RS.Fields("codgestion") & " - " & RS.Fields("gestion")),mid(trim(RS.Fields("codgestion") & " - " & RS.Fields("gestion")),1,30) & "..."),"")%>';
				datos[tabla][<%=contador%>][12]='<%=iif(RS.Fields("codgestion")<>"",RS.Fields("codgestion") & " - " & RS.Fields("gestion"),"")%>';
				datos[tabla][<%=contador%>][13]='<%=RS.Fields("tipocontacto")%>';
				datos[tabla][<%=contador%>][14]='<%=RS.Fields("fechadatos")%>';
				datos[tabla][<%=contador%>][15]='<%=RS.Fields("fechagestion")%>';		
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
	<%end if%>
<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
<form name=formula method=post>
<table border=0 cellspacing=0 cellpadding=2 width=100%>
  <tr bgcolor="#007DC5">
    <td colspan="17" align="center" height="22">
      <font size=2 face=Arial color="#FFFFFF"><b>Gestión de Impagados</b></font>
    </td>
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>N°&nbsp;Contrato:</b></font></td>
    <td colspan="4"><input name="contrato" type=text maxlength=18 value="<%=contrato%>" style="font-size: x-small; width: 250px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Código&nbsp;Cliente:</b></font></td>
    <td colspan="4"><input name="codigocentral" type=text maxlength=8 value="<%=codigocentral%>" style="font-size: x-small; width: 250px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Gestor:</b></font></td>
    <td colspan="4">
		<select name="codgestor" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Gestor</option>
				<%
				sql = "select A.codusuario,A.nombres,A.apepaterno,A.apematerno from Usuario A inner join TipoUsuario B on A.codtipousuario=B.codtipousuario and B.descripcion='Gestor' and A.activo=1"
				consultar sql,RS
				Do While Not  RS.EOF
				nombregestor=iif(IsNull(RS.Fields("nombres")),"",RS.Fields("nombres")) & ", " & iif(IsNull(RS.Fields("apepaterno")),"",RS.Fields("apepaterno")) & " " & iif(IsNull(RS.Fields("apematerno")),"",RS.Fields("apematerno"))
				%>
				<option value="<%=RS.Fields("codusuario")%>" <%if codgestor<>"" then%><%if RS.fields("codusuario")=int(codgestor) then%> selected<%end if%><%else%><%if actualizarlista="" and RS.fields("codusuario")=int(session("codusuario")) then%> selected<%end if%><%end if%>><%=nombregestor%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
				%>			
		</select>			
	</td>
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Tipo&nbsp;Documento:</b></font></td>
    <td colspan="4">
			<select name="codtipodocumento" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Tipo Documento</option>
				<%
				sql = "select codtipodocumento,descripcion from TipoDocumento where activo=1 order by descripcion"
				consultar sql,RS
				Do While Not RS.EOF
				%>
					<option value="<%=RS.Fields("codtipodocumento")%>" <% if codtipodocumento<>"" then%><% if RS.fields("codtipodocumento")=int(codtipodocumento) then%> selected<%end if%><%end if%>><%=RS.Fields("CodTipoDocumento") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
				%>				
			</select>
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>N°&nbsp;Documento:</b></font></td>
    <td colspan="4"><input name="numdocumento" type=text maxlength=15 value="<%=numdocumento%>" style="font-size: x-small; width: 250px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Días&nbsp;de&nbsp;Atraso:</b></font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Del</b></font></td>
  <td><input name="diaatrasoini" type=text maxlength=10 value="<%=diaatrasoini%>" style="text-align: right;font-size: x-small; width: 50px;"></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>al</b></font></td>
    <td><input name="diaatrasofin" type=text maxlength=10 value="<%=diaatrasofin%>" style="text-align: right;font-size: x-small; width: 50px;"></td>
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Producto</b></font></td>
    <td colspan="4">
    <select name="codfamproducto" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Producto</option>
				<%
				sql = "Select codfamproducto,descripcion from FamProducto"
				consultar sql,RS
				Do While Not  RS.EOF
				%>
					<option value="<%=RS.Fields("codfamproducto")%>" <%if codfamproducto<>"" then%><% if RS.fields("codfamproducto")=int(codfamproducto) then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
				%>
		</select>
	</td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Agencia:</b></font></td>
    <td colspan="4">
		<%if codagenciausuario<>"" then%>
			<font size=1 face=Arial color=#00529B><%=AgenciaUsuario%></font><input type=hidden name="codagencia" value="<%=codagenciausuario%>">
		<%else%>
			<select name="codagencia" style="font-size: x-small; width: 250px;">
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
		<%end if%>
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Marca:</b></font></td>
    <td colspan="4">
			<select name="codmarca" style="font-size: x-small; width: 250px;">
				<option value="">Seleccione Marca</option>
				<%
				sql = "select codmarca,descripcion from Marca where activo=1 order by codmarca"
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
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Territorio:</b></font></td>
    <td colspan="4">
		<%if codoficinausuario<>"" then%>
			<font size=1 face=Arial color=#00529B><%=TerritorioUsuario%></font>
		<%else%>    
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
		<%end if%>
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Oficina:</b></font></td>
    <td colspan="4">
		<%if codoficinausuario<>"" then%>
			<font size=1 face=Arial color=#00529B><%=codoficinausuario & " - " & OficinaUsuario%></font><input type=hidden name="codoficina" value="<%=codoficinausuario%>">
		<%else%>        
			<select name = "codoficina" style="font-size: x-small; width: 250px;">
			<option value="">Seleccione Oficina</option>
			<%
				if codterritorio<>"" then
					sql = "select codoficina, descripcion from oficina where activo = 1 and codterritorio = " & codterritorio & " order by codoficina"
					consultar sql,RS
					Do While Not  RS.EOF
					%>
					<option value="<%=RS.Fields("codoficina")%>" <% if codoficina<>"" then%><% if RS.fields("codoficina")=codoficina then%> selected<%end if%><%end if%>><%=RS.Fields("CodOficina") & " - " & RS.Fields("Descripcion")%></option>
					<%
					RS.MoveNext
					loop
					RS.Close
				end if
			%>
			</select>
		<%end if%> 
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Tipo&nbsp;de&nbsp;Gestión:</b></font></td>
    <td colspan="4">
		<select name = "tipogestion" style="font-size: x-small; width: 250px;">
			<option value="">Seleccione Gestión</option>
			<option value="Preventiva" <% if tipogestion<>"" then%><%if tipogestion="Preventiva" then%> selected<%end if%><%end if%>>Preventiva</option>
			<option value="Impagado" <% if tipogestion<>"" then%><%if tipogestion="Impagado" then%> selected<%end if%><%end if%>>Impagado</option>
		</select>    
    </td>
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Fecha&nbsp;de&nbsp;Asignación:</b></font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Del</b></font></td>
    <td width="78"><input name="fechaasigini"  id="sel0" readonly  type=text maxlength=10 size=10 value="<%if IsDate(fechaasigini) then%><%=fechaasigini%><%else%><%=maxfechagestion%><%end if%>"  style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel0', '%d/%m/%Y');"></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>al</b></font></td>
    <td>
		<input name="fechaasigfin" id="sel3" type=text maxlength=10 readonly value="<%if IsDate(fechaasigfin) then%><%=fechaasigfin%><%else%><%=maxfechagestion%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel3', '%d/%m/%Y');">
	<!--<input type="text" name="date3" id="sel3" size="10"><input type="image" value=" ... " onclick="return showCalendar('sel3', '%d/%m/%Y');">-->
	</td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Mejor&nbsp;Gestión:</b></font></td>
    <td colspan=4>
		<select name = "codgestion" style="font-size: x-small; width: 250px;">
			<option value="">Seleccione Mejor Gestión</option>
			<option value="*" <%if codgestion="*" then%> selected<%end if%>>Sin Gestión</option>
			<option value="**" <%if codgestion="**" then%> selected<%end if%>>Todas las Gestionadas</option>
			<%
				sql = "select codgestion, descripcion from Gestion where activo=1 order by codgestion"
				consultar sql,RS
				Do While Not RS.EOF
				%>
				<option value="<%=RS.Fields("codgestion")%>" <% if codgestion<>"" then%><% if RS.fields("codgestion")=codgestion then%> selected<%end if%><%end if%>><%=RS.Fields("codgestion") & " - " & RS.Fields("Descripcion")%></option>
				<%
				RS.MoveNext
				loop
				RS.Close
			%>
		</select>   
	</td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Fecha&nbsp;de&nbsp;Promesa:</b></font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Del</b></font></td>
    <td width="78"><input name="fechapromesaini" type=text maxlength=10 id="sel4"  value="<%if IsDate(fechapromesaini) then%><%=fechapromesaini%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel4', '%d/%m/%Y');"></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>al</b></font></td>
    <td><input name="fechapromesafin" type=text maxlength=10 id="sel5" value="<%if IsDate(fechapromesafin) then%><%=fechapromesafin%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel5', '%d/%m/%Y');"></td>    
  </tr>
  <tr>
  <td colspan="17">
		 <%if contador=0 then%>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=1 face=Arial color=#00529B><b>Impagados (0) - No hay registros.</b></font>&nbsp;<a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
			</tr>
			</table>
		<%else		
		%>
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=1 face=Arial color=#00529B><b>Impagados (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=right width=180><font size=1 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
			</tr>	
			</table>
			<div id="tabla0"> 
			</div>
		<%end if%>
        <input type="hidden" name="buscando" value="">				
		<input type="hidden" name="actualizarlista" value="<%=actualizarlista%>">
		<input type="hidden" name="nuevabusqueda" value="">
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="pag" value="<%=pag%>">
		<input type="hidden" name="contrato1" value="<%=obtener("contrato1")%>">
		<input type="hidden" name="codigocentral1" value="<%=obtener("codigocentral1")%>">
		<input type="hidden" name="codgestor1" value="<%=obtener("codgestor1")%>">
		<input type="hidden" name="codtipodocumento1" value="<%=obtener("codtipodocumento1")%>">
		<input type="hidden" name="numdocumento1" value="<%=obtener("numdocumento1")%>">
		<input type="hidden" name="diaatrasoini1" value="<%=obtener("diaatrasoini1")%>">
		<input type="hidden" name="diaatrasofin1" value="<%=obtener("diaatrasofin1")%>">
		<input type="hidden" name="codfamproducto1" value="<%=obtener("codfamproducto1")%>">
		<input type="hidden" name="codagencia1" value="<%=obtener("codagencia1")%>">
		<input type="hidden" name="codmarca1" value="<%=obtener("codmarca1")%>">
		<input type="hidden" name="codterritorio1" value="<%=obtener("codterritorio1")%>">
		<input type="hidden" name="codoficina1" value="<%=obtener("codoficina1")%>">
		<input type="hidden" name="tipogestion1" value="<%=obtener("tipogestion1")%>">
		<input type="hidden" name="fechaasigini1" value="<%=obtener("fechaasigini1")%>">
		<input type="hidden" name="fechaasigfin1" value="<%=obtener("fechaasigfin1")%>">
		<input type="hidden" name="fechapromesaini1" value="<%=obtener("fechapromesaini1")%>">
		<input type="hidden" name="fechapromesafin1" value="<%=obtener("fechapromesafin1")%>">
		<input type="hidden" name="codgestion1" value="<%=obtener("codgestion1")%>">
					
		<%if actualizarlista<>"" then%>
		<script language="javascript">
			inicio();
		</script>					
		<%end if%>
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>		
	
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
					consulta_exp="select 'Código','N° Contrato','Nombre','Teléfono','Producto','Subprod','Días Atraso','Moneda','Deuda','Agencia','F.Asig','F.Gest','Mejor Gestión','Tipo Contacto'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					
					
					
					
                    sql="select OBJECT_ID('tempdb.dbo.##User" & session("codusuariodif") & "')"
                    ''response.write sql
		            consultar sql,RS	
                    if obtener("nuevabusqueda")<>"" or IsNull(RS.Fields(0)) then
					    consulta_exp="select char(39) + A.codigocentral,char(39) + A.contrato,A.nombre,char(39) + CASE WHEN A.tipofono1='C' THEN A.fono1 ELSE CASE WHEN rtrim(ltrim(A.prefijo1))='001' THEN CASE WHEN rtrim(ltrim(A.extension1))<>'0000' and rtrim(ltrim(A.extension1))<>'' THEN A.fono1 + '.Ext.' + A.extension1 ELSE A.fono1 END ELSE CASE WHEN rtrim(ltrim(A.extension1))<>'0000' and rtrim(ltrim(A.extension1))<>'' THEN A.prefijo1 + A.fono1 + '.Ext.' + A.extension1 ELSE A.prefijo1 + A.fono1 END END END as telefono,A.producto,A.subproducto,A.diasvencimiento,A.divisa,A.saldohoy,A.Agencia,A.fechagestion,A.fhgestionado,A.codgestion + ' - ' + A.Gestion as Gestion,A.TipoContacto " & _
								    "from CobranzaCM.dbo." & vistabusqueda & " A " & replace(filtrobuscador,"from ","from CobranzaCM.dbo.") & " order by A.nro"								 
					    sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					    conn.execute sql                    
					else
					    consulta_exp="select char(39) + A.codigocentral,char(39) + A.contrato,A.nombre,char(39) + CASE WHEN A.tipofono1='C' THEN A.fono1 ELSE CASE WHEN rtrim(ltrim(A.prefijo1))='001' THEN CASE WHEN rtrim(ltrim(A.extension1))<>'0000' and rtrim(ltrim(A.extension1))<>'' THEN A.fono1 + '.Ext.' + A.extension1 ELSE A.fono1 END ELSE CASE WHEN rtrim(ltrim(A.extension1))<>'0000' and rtrim(ltrim(A.extension1))<>'' THEN A.prefijo1 + A.fono1 + '.Ext.' + A.extension1 ELSE A.prefijo1 + A.fono1 END END END as telefono,A.producto,A.subproducto,A.diasvencimiento,A.divisa,A.saldohoy,A.Agencia,A.fechagestion,A.fhgestionado,A.codgestion + ' - ' + A.Gestion as Gestion,A.TipoContacto " & _
								    "from tempdb.dbo.##User" & session("codusuariodif") & " A order by A.nro"								 
					    sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					    conn.execute sql                    					
                    end if
                    RS.close 
					
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
<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>
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
	window.open("index.html","_top");
</script>
<%
end if
%>

