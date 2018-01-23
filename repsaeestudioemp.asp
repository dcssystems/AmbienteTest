<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repsaeestudioemp.asp") then

	    fechacierre=obtener("fechacierre")
	    ''if not IsDate(fechacierre) then
	    if fechacierre="" then
            ''sql="select max(fechadatos) as fechadatos from FB.CentroCobranzas.dbo.PD_Detalle_Casos_JUD"
            sql="select max(fechadatos) as fechadatos from FB.CentroCobranzas.dbo.pd_casos_jud"
            consultar sql,RS	
            fechacierre=RS.Fields("fechadatos")
            RS.Close
		end if
		
		
	
		sql="select A.Plaza, " & _
		    "A.CODDATO, " & _
		            "IsNull(A.nomb_estu,'NO IDENTIFICADO') as ESTUDIO, " & _
		            "SUM(A.JUD) as MontoSoles, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC' THEN A.JUD ELSE 0 END) as MontoSolesBEC, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC-RED' THEN A.JUD ELSE 0 END) as MontoSolesBECRED, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PROMOTORES' THEN A.JUD ELSE 0 END) as MontoSolesPROMOTORES, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC' THEN 1 ELSE 0 END) as CasosBEC, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC-RED' THEN 1 ELSE 0 END) as CasosBECRED, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PROMOTORES' THEN 1 ELSE 0 END) as CasosPROMOTORES, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFaseBEC, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC-RED' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFaseBECRED, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PROMOTORES' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFasePROMOTORES, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFaseBEC, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='BEC-RED' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFaseBECRED, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PROMOTORES' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFasePROMOTORES " & _
            "from " & _
	            "(select A.CONTRATO,A.Plaza,A.CODDATO,A.nomb_estu,A.SEGMENTO_RIESGO,A.JUD,MAX(A.CODIGO) as CODIGO " & _
	            "from FB.CentroCobranzas.dbo.PD_Detalle_Casos_JUD A " & _
	            "where A.SEGMENTO_RIESGO in ('BEC','BEC-RED','PROMOTORES') and A.FASIGNA is not null and A.FECHADATOS='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'" & _
	            "group by A.CONTRATO,A.Plaza,A.CODDATO,A.nomb_estu,SEGMENTO_RIESGO,A.JUD) A " & _
            "group by A.Plaza,A.CODDATO,IsNull(A.nomb_estu,'NO IDENTIFICADO') " & _
            "order by A.Plaza,MontoSoles desc"
            ''"order by A.Plaza,IsNull(A.nomb_estu,'NO IDENTIFICADO')"
            ''"where A.FASIGNA<=DATEADD(m,-2,getdate()) " & _
        ''response.Write sql
		consultar sql,RS
		
		
		SUMA_MontoSoles_LIMA_BEC=0
		SUMA_MontoSoles_LIMA_BECRED=0
		SUMA_MontoSoles_LIMA_PROMOTORES=0
		SUMA_Cantidad_LIMA_BEC=0
		SUMA_Cantidad_LIMA_BECRED=0
		SUMA_Cantidad_LIMA_PROMOTORES=0
		SUMA_MontoSoles_LIMA_BEC_CF=0
		SUMA_MontoSoles_LIMA_BECRED_CF=0
		SUMA_MontoSoles_LIMA_PROMOTORES_CF=0
		SUMA_Cantidad_LIMA_BEC_CF=0
		SUMA_Cantidad_LIMA_BECRED_CF=0
		SUMA_Cantidad_LIMA_PROMOTORES_CF=0
		
		SUMA_MontoSoles_PROV_BEC=0
		SUMA_MontoSoles_PROV_BECRED=0
		SUMA_MontoSoles_PROV_PROMOTORES=0
		SUMA_Cantidad_PROV_BEC=0
		SUMA_Cantidad_PROV_BECRED=0
		SUMA_Cantidad_PROV_PROMOTORES=0
		SUMA_MontoSoles_PROV_BEC_CF=0
		SUMA_MontoSoles_PROV_BECRED_CF=0
		SUMA_MontoSoles_PROV_PROMOTORES_CF=0
		SUMA_Cantidad_PROV_BEC_CF=0
		SUMA_Cantidad_PROV_BECRED_CF=0
		SUMA_Cantidad_PROV_PROMOTORES_CF=0
		
		SUMA_MontoSoles_BEC=0
		SUMA_Cantidad_BEC=0
		SUMA_MontoSoles_BECRED=0
		SUMA_Cantidad_BECRED=0
		SUMA_MontoSoles_PROMOTORES=0
		SUMA_Cantidad_PROMOTORES=0		
		SUMA_MontoSoles_BEC_CF=0
		SUMA_Cantidad_BEC_CF=0
		SUMA_MontoSoles_BECRED_CF=0
		SUMA_Cantidad_BECRED_CF=0
		SUMA_MontoSoles_PROMOTORES_CF=0
		SUMA_Cantidad_PROMOTORES_CF=0

								
        RS.Filter=" Plaza='LIMA' "
        Do While not RS.EOF
		    SUMA_MontoSoles_LIMA_BEC=SUMA_MontoSoles_LIMA_BEC + RS.Fields("MontoSolesBEC")
		    SUMA_Cantidad_LIMA_BEC=SUMA_Cantidad_LIMA_BEC + RS.Fields("CasosBEC")
		    SUMA_MontoSoles_LIMA_BEC_CF=SUMA_MontoSoles_LIMA_BEC_CF + RS.Fields("MontoSolesConFaseBEC")
		    SUMA_Cantidad_LIMA_BEC_CF=SUMA_Cantidad_LIMA_BEC_CF + RS.Fields("CasosConFaseBEC")
		    SUMA_MontoSoles_LIMA_BECRED=SUMA_MontoSoles_LIMA_BECRED + RS.Fields("MontoSolesBECRED")
		    SUMA_Cantidad_LIMA_BECRED=SUMA_Cantidad_LIMA_BECRED + RS.Fields("CasosBECRED")
		    SUMA_MontoSoles_LIMA_BECRED_CF=SUMA_MontoSoles_LIMA_BECRED_CF + RS.Fields("MontoSolesConFaseBECRED")
		    SUMA_Cantidad_LIMA_BECRED_CF=SUMA_Cantidad_LIMA_BECRED_CF + RS.Fields("CasosConFaseBECRED")
		    SUMA_MontoSoles_LIMA_PROMOTORES=SUMA_MontoSoles_LIMA_PROMOTORES + RS.Fields("MontoSolesPROMOTORES")
		    SUMA_Cantidad_LIMA_PROMOTORES=SUMA_Cantidad_LIMA_PROMOTORES + RS.Fields("CasosPROMOTORES")
		    SUMA_MontoSoles_LIMA_PROMOTORES_CF=SUMA_MontoSoles_LIMA_PROMOTORES_CF + RS.Fields("MontoSolesConFasePROMOTORES")
		    SUMA_Cantidad_LIMA_PROMOTORES_CF=SUMA_Cantidad_LIMA_PROMOTORES_CF + RS.Fields("CasosConFasePROMOTORES")		    
        RS.MoveNext
        Loop 
        RS.Filter=""
        RS.Filter=" Plaza='PROVINCIA' "
        Do While not RS.EOF
		    SUMA_MontoSoles_PROV_BEC=SUMA_MontoSoles_PROV_BEC + RS.Fields("MontoSolesBEC")
		    SUMA_Cantidad_PROV_BEC=SUMA_Cantidad_PROV_BEC + RS.Fields("CasosBEC")
		    SUMA_MontoSoles_PROV_BEC_CF=SUMA_MontoSoles_PROV_BEC_CF + RS.Fields("MontoSolesConFaseBEC")
		    SUMA_Cantidad_PROV_BEC_CF=SUMA_Cantidad_PROV_BEC_CF + RS.Fields("CasosConFaseBEC")
		    SUMA_MontoSoles_PROV_BECRED=SUMA_MontoSoles_PROV_BECRED + RS.Fields("MontoSolesBECRED")
		    SUMA_Cantidad_PROV_BECRED=SUMA_Cantidad_PROV_BECRED + RS.Fields("CasosBECRED")
		    SUMA_MontoSoles_PROV_BECRED_CF=SUMA_MontoSoles_PROV_BECRED_CF + RS.Fields("MontoSolesConFaseBECRED")
		    SUMA_Cantidad_PROV_BECRED_CF=SUMA_Cantidad_PROV_BECRED_CF + RS.Fields("CasosConFaseBECRED")
		    SUMA_MontoSoles_PROV_PROMOTORES=SUMA_MontoSoles_PROV_PROMOTORES + RS.Fields("MontoSolesPROMOTORES")
		    SUMA_Cantidad_PROV_PROMOTORES=SUMA_Cantidad_PROV_PROMOTORES + RS.Fields("CasosPROMOTORES")
		    SUMA_MontoSoles_PROV_PROMOTORES_CF=SUMA_MontoSoles_PROV_PROMOTORES_CF + RS.Fields("MontoSolesConFasePROMOTORES")
		    SUMA_Cantidad_PROV_PROMOTORES_CF=SUMA_Cantidad_PROV_PROMOTORES_CF + RS.Fields("CasosConFasePROMOTORES")		    
        RS.MoveNeXt
        Loop    

		SUMA_MontoSoles_BEC=SUMA_MontoSoles_LIMA_BEC + SUMA_MontoSoles_PROV_BEC
		SUMA_Cantidad_BEC=SUMA_Cantidad_LIMA_BEC + SUMA_Cantidad_PROV_BEC
		SUMA_MontoSoles_BECRED=SUMA_MontoSoles_LIMA_BECRED + SUMA_MontoSoles_PROV_BECRED
		SUMA_Cantidad_BECRED=SUMA_Cantidad_LIMA_BECRED + SUMA_Cantidad_PROV_BECRED
		SUMA_MontoSoles_PROMOTORES=SUMA_MontoSoles_LIMA_PROMOTORES + SUMA_MontoSoles_PROV_PROMOTORES
		SUMA_Cantidad_PROMOTORES=SUMA_Cantidad_LIMA_PROMOTORES + SUMA_Cantidad_PROV_PROMOTORES		
		SUMA_MontoSoles_BEC_CF=SUMA_MontoSoles_LIMA_BEC_CF + SUMA_MontoSoles_PROV_BEC_CF
		SUMA_Cantidad_BEC_CF=SUMA_Cantidad_LIMA_BEC_CF + SUMA_Cantidad_PROV_BEC_CF
		SUMA_MontoSoles_BECRED_CF=SUMA_MontoSoles_LIMA_BECRED_CF + SUMA_MontoSoles_PROV_BECRED_CF
		SUMA_Cantidad_BECRED_CF=SUMA_Cantidad_LIMA_BECRED_CF + SUMA_Cantidad_PROV_BECRED_CF
		SUMA_MontoSoles_PROMOTORES_CF=SUMA_MontoSoles_LIMA_PROMOTORES_CF + SUMA_MontoSoles_PROV_PROMOTORES_CF
		SUMA_Cantidad_PROMOTORES_CF=SUMA_Cantidad_LIMA_PROMOTORES_CF + SUMA_Cantidad_PROV_PROMOTORES_CF		
		%>
		
		
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
			<title>SAE por Estudios - Empresas</title>
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
			            //alert(document.formula.fechacierre.value);
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
            var columnas=22; //CANTIDAD DE COLUMNAS//

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
						<font size=4 color=#483d8b face=Arial><b>Alertas por Estudio - Empresas al <%=fechacierre%></b></font></td>
                        
					</tr>
					</table>
					<div id="cabtabla_rep" style="overflow:auto; height:auto; padding:0;"><!--margin-right: 17px;">-->
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#007DC5" rowspan=3>
							<font size=1 color=#FFFFFF face=Arial><b>ESTUDIOS</b></font>
						</td>
						<td bgcolor="#007DC5" colspan=7 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>BEC</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=7 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>BEC-RED</b></font>
						</td>
						<td bgcolor="#007DC5" colspan=7 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>PROMOTORES</b></font>
						</td>						
						<!--<td bgcolor="#FFFFFF" rowspan=3 width=6>
						    <font size=1 color=#FFFFFF face=Arial>&nbsp;</font>
						</td>-->												
					</tr>
					<tr>
						<td bgcolor="#007DC5" colspan=3 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Asignado</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=4 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Con Fase</b></font>
						</td>												
						<td bgcolor="#007DC5" colspan=3 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Asignado</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=4 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>Con Fase</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=3 align="center">
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
					</tr>		
					</table>
					</div>
					<div id="dettabla_rep" style="overflow:auto; height:80%; padding:0;">
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>Total</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_BEC/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_BEC%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_BEC_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_BEC_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"sem_verde",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_BECRED/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_BECRED%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_BECRED_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_BECRED_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_BECRED_CF*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=BEC-RED&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_BECRED_CF*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1)>=90,"sem_verde",iif(SUMA_MontoSoles_BECRED_CF*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROMOTORES/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROMOTORES%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROMOTORES_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROMOTORES_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=4%>
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=PROMOTORES&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_BEC_CF*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
					</tr>					
					<!--LIMA-->
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Lima</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BEC/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_BEC%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BEC*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BEC_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_BEC_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_BEC_CF*100/iif(SUMA_MontoSoles_LIMA_BEC>0,SUMA_MontoSoles_LIMA_BEC,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BECRED/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_BECRED%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BECRED*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BECRED_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_BECRED_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=BEC-RED&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_BECRED_CF*100/iif(SUMA_MontoSoles_LIMA_BECRED>0,SUMA_MontoSoles_LIMA_BECRED,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PROMOTORES/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PROMOTORES%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PROMOTORES*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PROMOTORES_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PROMOTORES_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=PROMOTORES&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_PROMOTORES_CF*100/iif(SUMA_MontoSoles_LIMA_PROMOTORES>0,SUMA_MontoSoles_LIMA_PROMOTORES,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>						
					</tr>	
					<%
                        RS.Filter=""
                        RS.Filter=" Plaza='LIMA' "
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
		                        <a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=&confase="><font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESTUDIO")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBEC")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosBEC")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBEC")*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBEC")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFaseBEC")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBECRED")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosBECRED")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBECRED")*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBECRED")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFaseBECRED")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=BEC-RED&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROMOTORES")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPROMOTORES")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROMOTORES")*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePROMOTORES")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePROMOTORES")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=PROMOTORES&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>	                        
                        </tr>	                        
                        <%
                        RS.MoveNext
                        Loop 					
					%>			
					<!--PROVINCIA-->	
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Provincia</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BEC/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_BEC%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BEC*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BEC_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_BEC_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_BEC_CF*100/iif(SUMA_MontoSoles_PROV_BEC>0,SUMA_MontoSoles_PROV_BEC,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BECRED/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_BECRED%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BECRED*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BECRED_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_BECRED_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=BEC-RED&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_BECRED_CF*100/iif(SUMA_MontoSoles_PROV_BECRED>0,SUMA_MontoSoles_PROV_BECRED,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PROMOTORES/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PROMOTORES%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PROMOTORES*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PROMOTORES_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PROMOTORES_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=PROMOTORES&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_PROMOTORES_CF*100/iif(SUMA_MontoSoles_PROV_PROMOTORES>0,SUMA_MontoSoles_PROV_PROMOTORES,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>						
					</tr>	
                        <%
                        RS.Filter=""
                        RS.Filter=" Plaza='PROVINCIA' "
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
		                        <a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=&confase="><font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESTUDIO")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBEC")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosBEC")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBEC")*100/iif(SUMA_MontoSoles_BEC>0,SUMA_MontoSoles_BEC,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBEC")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFaseBEC")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosBEC")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBEC")*100/iif(RS.Fields("MontoSolesBEC")>0,RS.Fields("MontoSolesBEC"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a><%end if%></td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBECRED")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosBECRED")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesBECRED")*100/iif(SUMA_MontoSoles_BECRED>0,SUMA_MontoSoles_BECRED,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBECRED")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFaseBECRED")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=BEC-RED&confase=1"><%if RS.Fields("CasosBECRED")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFaseBECRED")*100/iif(RS.Fields("MontoSolesBECRED")>0,RS.Fields("MontoSolesBECRED"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROMOTORES")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPROMOTORES")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPROMOTORES")*100/iif(SUMA_MontoSoles_PROMOTORES>0,SUMA_MontoSoles_PROMOTORES,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePROMOTORES")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePROMOTORES")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosempdet.asp?paginapadre=repsaeestudioemp.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=PROMOTORES&confase=1"><%if RS.Fields("CasosPROMOTORES")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePROMOTORES")*100/iif(RS.Fields("MontoSolesPROMOTORES")>0,RS.Fields("MontoSolesPROMOTORES"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>	                        
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

