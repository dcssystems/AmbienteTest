<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repsaegestor.asp") then
	
		fechacierre=obtener("fechacierre")
	    ''if not IsDate(fechacierre) then
	    if fechacierre="" then
            sql="select max(fechadatos) as fechadatos from FB.CentroCobranzas.dbo.pd_casos_jud"
            consultar sql,RS	
            fechacierre=RS.Fields("fechadatos")
            RS.Close
		end if
		
	
		sql="select A.Plaza, " & _
		            "A.ESPECIALISTA as CODGESTOR, " & _
		            "A.ESPECIALISTA, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' THEN A.JUD ELSE 0 END) as MontoSolesPART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' and A.FASIGNA is not null THEN A.JUD ELSE 0 END) as MontoSolesPARTAsig, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' THEN A.JUD ELSE 0 END) as MontoSolesPYME, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' and A.FASIGNA is not null THEN A.JUD ELSE 0 END) as MontoSolesPYMEAsig, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' THEN 1 ELSE 0 END) as CasosPART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' and A.FASIGNA is not null THEN 1 ELSE 0 END) as CasosPARTAsig, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' THEN 1 ELSE 0 END) as CasosPYME, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' and A.FASIGNA is not null THEN 1 ELSE 0 END) as CasosPYMEAsig, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFasePART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFasePYME, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFasePART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFasePYME " & _
            "from " & _
	            "(select A.CONTRATO,A.Plaza,A.ESPECIALISTA,A.SEGMENTO_RIESGO,A.JUD,A.FASIGNA,MAX(A.CODIGO) as CODIGO " & _
	            "from FB.CentroCobranzas.dbo.PD_Detalle_Casos_JUD A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') " & _
	            "group by A.CONTRATO,A.Plaza,A.ESPECIALISTA,A.SEGMENTO_RIESGO,A.JUD,A.FASIGNA) A " & _
            "group by A.Plaza,A.ESPECIALISTA " & _
            "order by A.Plaza,MontoSolesPART desc,MontoSolesPYME desc"
            ''"order by A.Plaza,IsNull(A.nomb_estu,'NO IDENTIFICADO')"
            ''"where A.FASIGNA<=DATEADD(m,-2,getdate()) " & _


		sql="select A.SEGMENTO_RIESGO, " & _
		            "A.ESPECIALISTA as CODGESTOR, " & _
		            "A.ESPECIALISTA, " & _
		            "SUM(A.JUD) as MontoSoles, " & _
		            "SUM(CASE WHEN A.Plaza='LIMA' THEN A.JUD ELSE 0 END) as MontoSolesLIMA, " & _
		            "SUM(CASE WHEN A.Plaza='LIMA' and A.FASIGNA is not null THEN A.JUD ELSE 0 END) as MontoSolesLIMAAsig, " & _
		            "SUM(CASE WHEN A.Plaza='PROVINCIA' THEN A.JUD ELSE 0 END) as MontoSolesPROVINCIA, " & _
		            "SUM(CASE WHEN A.Plaza='PROVINCIA' and A.FASIGNA is not null THEN A.JUD ELSE 0 END) as MontoSolesPROVINCIAAsig, " & _
		            "SUM(CASE WHEN A.Plaza='LIMA' THEN 1 ELSE 0 END) as CasosLIMA, " & _
		            "SUM(CASE WHEN A.Plaza='LIMA' and A.FASIGNA is not null THEN 1 ELSE 0 END) as CasosLIMAAsig, " & _
		            "SUM(CASE WHEN A.Plaza='PROVINCIA' THEN 1 ELSE 0 END) as CasosPROVINCIA, " & _
		            "SUM(CASE WHEN A.Plaza='PROVINCIA' and A.FASIGNA is not null THEN 1 ELSE 0 END) as CasosPROVINCIAAsig, " & _
		            "SUM(CASE WHEN A.Plaza='LIMA' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFaseLIMA, " & _
		            "SUM(CASE WHEN A.Plaza='PROVINCIA' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFasePROVINCIA, " & _
		            "SUM(CASE WHEN A.Plaza='LIMA' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFaseLIMA, " & _
		            "SUM(CASE WHEN A.Plaza='PROVINCIA' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFasePROVINCIA " & _
            "from " & _
	            "(select A.CONTRATO,A.ESPECIALISTA,A.Plaza,A.SEGMENTO_RIESGO,A.JUD,A.FASIGNA,MAX(A.CODIGO) as CODIGO " & _
	            "from FB.CentroCobranzas.dbo.PD_Detalle_Casos_JUD A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and A.FECHADATOS='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'" & _
	            "group by A.CONTRATO,A.ESPECIALISTA,A.Plaza,A.SEGMENTO_RIESGO,A.JUD,A.FASIGNA) A " & _
            "group by A.SEGMENTO_RIESGO,A.ESPECIALISTA " & _
            "order by A.SEGMENTO_RIESGO,MontoSoles desc"
            ''"order by A.Plaza,IsNull(A.nomb_estu,'NO IDENTIFICADO')"



		consultar sql,RS
		
		
		SUMA_MontoSoles_PART_LIMA=0
		SUMA_MontoSoles_PYME_LIMA=0
		SUMA_Cantidad_PART_LIMA=0
		SUMA_Cantidad_PYME_LIMA=0
		SUMA_MontoSoles_PART_LIMAAsig=0
		SUMA_MontoSoles_PYME_LIMAAsig=0
		SUMA_Cantidad_PART_LIMAAsig=0
		SUMA_Cantidad_PYME_LIMAAsig=0		
		SUMA_MontoSoles_PART_LIMA_CF=0
		SUMA_MontoSoles_PYME_LIMA_CF=0
		SUMA_Cantidad_PART_LIMA_CF=0
		SUMA_Cantidad_PYME_LIMA_CF=0
		
		SUMA_MontoSoles_PART_PROV=0
		SUMA_MontoSoles_PYME_PROV=0
		SUMA_Cantidad_PART_PROV=0
		SUMA_Cantidad_PYME_PROV=0
		SUMA_MontoSoles_PART_PROVAsig=0
		SUMA_MontoSoles_PYME_PROVAsig=0
		SUMA_Cantidad_PART_PROVAsig=0
		SUMA_Cantidad_PYME_PROVAsig=0		
		SUMA_MontoSoles_PART_PROV_CF=0
		SUMA_MontoSoles_PYME_PROV_CF=0
		SUMA_Cantidad_PART_PROV_CF=0
		SUMA_Cantidad_PYME_PROV_CF=0
		
		SUMA_MontoSoles_LIMA=0
		SUMA_Cantidad_LIMA=0
		SUMA_MontoSoles_LIMAAsig=0
		SUMA_Cantidad_LIMAAsig=0		
		SUMA_MontoSoles_PROV=0
		SUMA_Cantidad_PROV=0
		SUMA_MontoSoles_PROVAsig=0
		SUMA_Cantidad_PROVAsig=0
		
		SUMA_MontoSoles_LIMA_CF=0
		SUMA_Cantidad_LIMA_CF=0
		SUMA_MontoSoles_PROV_CF=0
		SUMA_Cantidad_PROV_CF=0		
										
        RS.Filter=" SEGMENTO_RIESGO='PARTICULARES' "
        Do While not RS.EOF
		    SUMA_MontoSoles_PART_LIMA=SUMA_MontoSoles_PART_LIMA + RS.Fields("MontoSolesLIMA")
		    SUMA_Cantidad_PART_LIMA=SUMA_Cantidad_PART_LIMA + RS.Fields("CasosLIMA")
		    SUMA_MontoSoles_PART_LIMAAsig=SUMA_MontoSoles_PART_LIMAAsig + RS.Fields("MontoSolesLIMAAsig")
		    SUMA_Cantidad_PART_LIMAAsig=SUMA_Cantidad_PART_LIMAAsig + RS.Fields("CasosLIMAAsig")		    
		    SUMA_MontoSoles_PART_LIMA_CF=SUMA_MontoSoles_PART_LIMA_CF + RS.Fields("MontoSolesConFaseLIMA")
		    SUMA_Cantidad_PART_LIMA_CF=SUMA_Cantidad_PART_LIMA_CF + RS.Fields("CasosConFaseLIMA")
		    
		    SUMA_MontoSoles_PART_PROV=SUMA_MontoSoles_PART_PROV + RS.Fields("MontoSolesPROVINCIA")
		    SUMA_Cantidad_PART_PROV=SUMA_Cantidad_PART_PROV + RS.Fields("CasosPROVINCIA")
		    SUMA_MontoSoles_PART_PROVAsig=SUMA_MontoSoles_PART_PROVAsig + RS.Fields("MontoSolesPROVINCIAAsig")
		    SUMA_Cantidad_PART_PROVAsig=SUMA_Cantidad_PART_PROVAsig + RS.Fields("CasosPROVINCIAAsig")		    
		    SUMA_MontoSoles_PART_PROV_CF=SUMA_MontoSoles_PART_PROV_CF + RS.Fields("MontoSolesConFasePROVINCIA")
		    SUMA_Cantidad_PART_PROV_CF=SUMA_Cantidad_PART_PROV_CF + RS.Fields("CasosConFasePROVINCIA")		    
        RS.MoveNeXt
        Loop 
        RS.Filter=""
        RS.Filter=" SEGMENTO_RIESGO='PYME' "
        Do While not RS.EOF
		    SUMA_MontoSoles_PYME_LIMA=SUMA_MontoSoles_PYME_LIMA + RS.Fields("MontoSolesLIMA")
		    SUMA_Cantidad_PYME_LIMA=SUMA_Cantidad_PYME_LIMA + RS.Fields("CasosLIMA")
		    SUMA_MontoSoles_PYME_LIMAAsig=SUMA_MontoSoles_PYME_LIMAAsig + RS.Fields("MontoSolesLIMAAsig")
		    SUMA_Cantidad_PYME_LIMAAsig=SUMA_Cantidad_PYME_LIMAAsig + RS.Fields("CasosLIMAAsig")		    
		    SUMA_MontoSoles_PYME_LIMA_CF=SUMA_MontoSoles_PYME_LIMA_CF + RS.Fields("MontoSolesConFaseLIMA")
		    SUMA_Cantidad_PYME_LIMA_CF=SUMA_Cantidad_PYME_LIMA_CF + RS.Fields("CasosConFaseLIMA")
		    
		    SUMA_MontoSoles_PYME_PROV=SUMA_MontoSoles_PYME_PROV + RS.Fields("MontoSolesPROVINCIA")
		    SUMA_Cantidad_PYME_PROV=SUMA_Cantidad_PYME_PROV + RS.Fields("CasosPROVINCIA")
		    SUMA_MontoSoles_PYME_PROVAsig=SUMA_MontoSoles_PYME_PROVAsig + RS.Fields("MontoSolesPROVINCIAAsig")
		    SUMA_Cantidad_PYME_PROVAsig=SUMA_Cantidad_PYME_PROVAsig + RS.Fields("CasosPROVINCIAAsig")		    
		    SUMA_MontoSoles_PYME_PROV_CF=SUMA_MontoSoles_PYME_PROV_CF + RS.Fields("MontoSolesConFasePROVINCIA")
		    SUMA_Cantidad_PYME_PROV_CF=SUMA_Cantidad_PYME_PROV_CF + RS.Fields("CasosConFasePROVINCIA")		    
        RS.MoveNeXt
        Loop    

		SUMA_MontoSoles_LIMA=SUMA_MontoSoles_PART_LIMA + SUMA_MontoSoles_PYME_LIMA
		SUMA_Cantidad_LIMA=SUMA_Cantidad_PART_LIMA + SUMA_Cantidad_PYME_LIMA 
		SUMA_MontoSoles_PROV=SUMA_MontoSoles_PART_PROV + SUMA_MontoSoles_PYME_PROV
		SUMA_Cantidad_PROV=SUMA_Cantidad_PYME_PROV + SUMA_Cantidad_PYME_PROV

		SUMA_MontoSoles_LIMAAsig=SUMA_MontoSoles_PART_LIMAAsig + SUMA_MontoSoles_PYME_LIMAAsig
		SUMA_Cantidad_LIMAAsig=SUMA_Cantidad_PART_LIMAAsig + SUMA_Cantidad_PYME_LIMAAsig 
		SUMA_MontoSoles_PROVAsig=SUMA_MontoSoles_PART_PROVAsig + SUMA_MontoSoles_PYME_PROVAsig
		SUMA_Cantidad_PROVAsig=SUMA_Cantidad_PART_PROVAsig + SUMA_Cantidad_PYME_PROVAsig
		
		SUMA_MontoSoles_LIMA_CF=SUMA_MontoSoles_PART_LIMA_CF + SUMA_MontoSoles_PYME_LIMA_CF
		SUMA_Cantidad_LIMA_CF=SUMA_Cantidad_PART_LIMA_CF + SUMA_Cantidad_PYME_LIMA_CF
		SUMA_MontoSoles_PROV_CF=SUMA_MontoSoles_PART_PROV_CF + SUMA_MontoSoles_PYME_PROV_CF
		SUMA_Cantidad_PROV_CF=SUMA_Cantidad_PART_PROV_CF + SUMA_Cantidad_PYME_PROV_CF
		%>
		
		
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
			<title>SAE por Gestor</title>
			<script language=javascript>
			    var ventanaverimpagado;
			    function inicio() {
			        dibujarTabla(0);
			    }
			    function modificar(codcen, contr, fd, fg) {
			        ventanaverimpagado = window.open("verimpagado.asp?vistapadre=" + window.name + "&paginapadre=admimpagado.asp&codigocentral=" + codcen + "&contrato=" + contr + "&fechadatos=" + fd + "&fechagestion=" + fg, "VerImpagado" + codcen, "scrollbars=yes,scrolling=yes,top=" + ((screen.height) / 2 - 300) + ",height=600,width=" + (screen.width / 2 + 300) + ",left=" + (screen.width / 2 - 475) + ",resizable=yes");
			        ventanaverimpagado.focus();
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
			        document.formula.fechacierre.value = document.formula.fechacierrebusc.value;
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
				FONT-SIZE: 12px; COLOR: #483d8b; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
			}
			A:visited {
				TEXT-DECORATION: none; COLOR: #483d8b;
			}
			A:hover {
				COLOR: #483d8b; FONT-FACE:"Arial"; TEXT-DECORATION: none
			}			
            #encabezado{border:0}
            #encabezado th{border-width:1px}
            #datos{border:0}
            #datos td{border-width:1px}			
			</style>
			
            <SCRIPT language= "JavaScript">
            var ancho1,ancho2,i;
            var columnas=15; //CANTIDAD DE COLUMNAS//

            function ajustaCeldas()
            {
                for (i = 0; i < columnas; i++)
                {
                    ancho1=document.getElementById("encabezado").rows.item(0).cells.item(i).offsetWidth;
                    ancho2=document.getElementById("datos").rows.item(0).cells.item(i).offsetWidth;
                    if (ancho1 > ancho2)
                        {
                        document.getElementById("datos").rows.item(0).cells.item(i).width = ancho1 - 6;
                        }
                        else
                        {
                            document.getElementById("encabezado").rows.item(0).cells.item(i).width = ancho2 - 6;
                        }
                }
            }

            function cuadratabla() 
            {
                if (document.getElementById('dettabla_rep').scrollHeight > document.getElementById('dettabla_rep').clientHeight) { document.getElementById('cabtabla_rep').style['margin-right'] = '17px'; } else { document.getElementById('cabtabla_rep').style['margin-right'] = '0px'; }             
            }
            </SCRIPT>

            
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
		   
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF" onload="cuadratabla();" onresize="cuadratabla();">
			<form name=formula method=post>
					<table border=0 cellpadding=0 cellspacing=0 width=100%>
					<tr>
					    <td bgcolor="#F5F5F5" width="120">&nbsp;<input name="fechacierrebusc" id="fechacierrebusc" readonly type=text maxlength=10 size=10 value="<%=fechacierre%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('fechacierrebusc', '%d/%m/%Y');">&nbsp;&nbsp;<a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle height="18"></a></td>
						<td bgcolor="#F5F5F5" align="center">			
						<font size=4 color=#483d8b face=Arial><b>Alertas por Gestor al <%=fechacierre%></b></font></td>
					</tr>
					</table>
					<div id="cabtabla_rep" style="overflow:auto; height:auto; padding:0;"><!--margin-right: 17px;">-->
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#007DC5" rowspan=3>
							<font size=1 color=#FFFFFF face=Arial><b>GESTOR</b></font>
						</td>
						<td bgcolor="#007DC5" colspan=11 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>LIMA</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=11 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>PROVINCIA</b></font>
						</td>
						<!--<td bgcolor="#FFFFFF" rowspan=3 width=6>
						    <font size=1 color=#FFFFFF face=Arial>&nbsp;</font>
						</td>-->												
					</tr>
					<tr>
						<td bgcolor="#007DC5" colspan=3 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Judicial</b></font>
						</td>						
						<td bgcolor="#007DC5" colspan=4 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Asignado</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=4 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Con Fase</b></font>
						</td>												
						<td bgcolor="#007DC5" colspan=3 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Judicial</b></font>
						</td>							
						<td bgcolor="#007DC5" colspan=4 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Asignado</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=4 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Con Fase</b></font>
						</td>																
					</tr>
					<tr>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" width=15>
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>						
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" width=15>
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" width=15>
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=4%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" width=15>
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>						
					</tr>		
					</table>
					</div>
					<div id="dettabla_rep" style="overflow:auto; height:80%; padding:0;">
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>Total</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
                        <td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMAAsig/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMAAsig%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_CF*100/iif(SUMA_MontoSoles_LIMAAsig>0,SUMA_MontoSoles_LIMAAsig,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROVAsig/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROVAsig%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROVAsig*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROVAsig*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROVAsig*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMAAsig*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_CF*100/iif(SUMA_MontoSoles_PROVAsig>0,SUMA_MontoSoles_PROVAsig,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
					</tr>					
					<!--PARTICULARES-->
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Particulares</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_LIMA/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_LIMA%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_LIMA*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_LIMAAsig/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_LIMAAsig%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase="><img src="imagenes/<%=iif(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_LIMAAsig*100/iif(SUMA_MontoSoles_PART_LIMA>0,SUMA_MontoSoles_PART_LIMA,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_LIMA_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_LIMA_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_LIMA_CF*100/iif(SUMA_MontoSoles_PART_LIMAAsig>0,SUMA_MontoSoles_PART_LIMAAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_PROV/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_PROV%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_PROV*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_PROVAsig/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_PROVAsig%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase="><img src="imagenes/<%=iif(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_PROVAsig*100/iif(SUMA_MontoSoles_PART_PROV>0,SUMA_MontoSoles_PART_PROV,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_PROV_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_PROV_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_PROV_CF*100/iif(SUMA_MontoSoles_PART_PROVAsig>0,SUMA_MontoSoles_PART_PROVAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
					</tr>	
					<%
                        RS.Filter=""
                        RS.Filter=" SEGMENTO_RIESGO='PARTICULARES' "
                        bgcolor="#FFFFFF"
                        Do While not RS.EOF
                            if bgcolor="#FFFFFF"then
                                bgcolor="#F5F5F5"
                            else
                                bgcolor="#FFFFFF"
                            end if
                        %>
                        <tr>
	                        <td bgcolor="<%=bgcolor%>">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESPECIALISTA")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesLIMA")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosLIMA")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesLIMA")*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesLIMAAsig")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosLIMAAsig")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase="><img src="imagenes/<%=iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a><%end if%></td>	                        	
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseLIMA")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFaseLIMA")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=PARTICULARES&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROVINCIA")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPROVINCIA")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROVINCIA")*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROVINCIAAsig")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPROVINCIAAsig")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase="><img src="imagenes/<%=iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>	                        
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePROVINCIA")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePROVINCIA")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PARTICULARES&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>
                        </tr>	                        
                        <%
                        RS.MoveNext
                        Loop 					
					%>			
					<!--PYME-->	
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Pyme</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_LIMA/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_LIMA%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_LIMA*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_LIMAAsig/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_LIMAAsig%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase="><img src="imagenes/<%=iif(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_LIMAAsig*100/iif(SUMA_MontoSoles_PYME_LIMA>0,SUMA_MontoSoles_PYME_LIMA,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_LIMA_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_LIMA_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_LIMA_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=lima&segmento=pyme&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PYME_LIMA_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PYME_LIMA_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PYME_LIMA_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_LIMA_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PYME_LIMA_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_LIMAAsig_CF*100/iif(SUMA_MontoSoles_PYME_LIMAAsig>0,SUMA_MontoSoles_PYME_LIMAAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_PROV/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_PROV%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_PROV*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_PROVAsig/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_PROVAsig%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase="><img src="imagenes/<%=iif(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_PROVAsig*100/iif(SUMA_MontoSoles_PYME_PROV>0,SUMA_MontoSoles_PYME_PROV,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_PROV_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_PROV_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=&plaza=provincia&segmento=pyme&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PYME_PROV_CF*100/iif(SUMA_MontoSoles_PYME_PROVAsig>0,SUMA_MontoSoles_PYME_PROVAsig,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
					</tr>	
                        <%
                        RS.Filter=""
                        RS.Filter=" SEGMENTO_RIESGO='PYME' "
                        bgcolor="#FFFFFF"
                        Do While not RS.EOF
                            if bgcolor="#FFFFFF"then
                                bgcolor="#F5F5F5"
                            else
                                bgcolor="#FFFFFF"
                            end if
                        %>
                        <tr>
	                        <td bgcolor="<%=bgcolor%>">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESPECIALISTA")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesLIMA")/1000,0)%><%end if%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=RS.Fields("CasosLIMA")%><%end if%></font></a>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesLIMA")*100/iif(SUMA_MontoSoles_LIMA>0,SUMA_MontoSoles_LIMA,1),2)%>%<%end if%></font></a>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesLIMAAsig")/1000,0)%><%end if%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=RS.Fields("CasosLIMAAsig")%><%end if%></font></a>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1),2)%>%<%end if%></font></a>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase="><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesLIMAAsig")*100/iif(RS.Fields("MontoSolesLIMA")>0,RS.Fields("MontoSolesLIMA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>	                        
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesConFaseLIMA")/1000,0)%><%end if%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=RS.Fields("CasosConFaseLIMA")%><%end if%></font></a>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1),2)%>%<%end if%></font></a>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=lima&segmento=pyme&confase=1"><%if RS.Fields("CasosLIMA")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseLIMA")*100/iif(RS.Fields("MontoSolesLIMAAsig")>0,RS.Fields("MontoSolesLIMAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesPROVINCIA")/1000,0)%><%end if%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=RS.Fields("CasosPROVINCIA")%><%end if%></font></a>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesPROVINCIA")*100/iif(SUMA_MontoSoles_PROV>0,SUMA_MontoSoles_PROV,1),2)%>%<%end if%></font></a>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesPROVINCIAAsig")/1000,0)%><%end if%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=RS.Fields("CasosPROVINCIAAsig")%><%end if%></font></a>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1),2)%>%<%end if%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase="><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesPROVINCIAAsig")*100/iif(RS.Fields("MontoSolesPROVINCIA")>0,RS.Fields("MontoSolesPROVINCIA"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>	                        	
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase=1"><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesConFasePROVINCIA")/1000,0)%><%end if%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase=1"><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=RS.Fields("CasosConFasePROVINCIA")%><%end if%></font></a>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase=1"><font size=1 color=#483d8b face=Arial><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><%=FormatNumber(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1),2)%>%<%end if%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaegestor.asp&asignado=S&fechacierre=<%=fechacierre%>&codgestor=<%=replace(RS.Fields("codgestor")," ","")%>&plaza=provincia&segmento=PYME&confase=1"><%if RS.Fields("CasosPROVINCIA")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROVINCIA")*100/iif(RS.Fields("MontoSolesPROVINCIAAsig")>0,RS.Fields("MontoSolesPROVINCIAAsig"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>
                        </tr>	                        
                        <%
                        RS.MoveNext
                        Loop 					
					%>		
                    <tr>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>						
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>				
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>		
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>						
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>				
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>						
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>				
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>		
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>						
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>				
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>						
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>				
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>						
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>				
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>
						<td bgcolor="#007DC5" align="right">
							<font size=1 color=#483d8b face=Arial>&nbsp;</font>
						</td>												
					</tr>																		
					</table>
					</div>
			<input type="hidden" name="expimp" value="">		
		        <input type="hidden" name="pag" value="<%=pag%>">
		        <input type="hidden" name="fechacierre" value="<%=obtener("fechacierre")%>">
			</form>							
			</body>
		</html>	
		
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display = "none";</script>	
		<%	
		Rs.Close	
		''end if
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
	//window.open("index.html","_top");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>

