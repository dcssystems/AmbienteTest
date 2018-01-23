<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repsaeestudio.asp") then

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
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' THEN A.JUD ELSE 0 END) as MontoSolesPART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' THEN A.JUD ELSE 0 END) as MontoSolesPYME, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' THEN 1 ELSE 0 END) as CasosPART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' THEN 1 ELSE 0 END) as CasosPYME, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFasePART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' AND A.CODIGO is not null THEN A.JUD ELSE 0 END) as MontoSolesConFasePYME, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PARTICULARES' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFasePART, " & _
		            "SUM(CASE WHEN A.SEGMENTO_RIESGO='PYME' AND A.CODIGO is not null THEN 1 ELSE 0 END) as CasosConFasePYME " & _
            "from " & _
	            "(select A.CONTRATO,A.Plaza,A.CODDATO,A.nomb_estu,A.SEGMENTO_RIESGO,A.JUD,MAX(A.CODIGO) as CODIGO " & _
	            "from FB.CentroCobranzas.dbo.PD_Detalle_Casos_JUD A " & _
	            "where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and A.FASIGNA is not null and A.FECHADATOS='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'" & _
	            "group by A.CONTRATO,A.Plaza,A.CODDATO,A.nomb_estu,SEGMENTO_RIESGO,A.JUD) A " & _
            "group by A.Plaza,A.CODDATO,IsNull(A.nomb_estu,'NO IDENTIFICADO') " & _
            "order by A.Plaza,MontoSoles desc"
            ''"order by A.Plaza,IsNull(A.nomb_estu,'NO IDENTIFICADO')"
            ''"where A.FASIGNA<=DATEADD(m,-2,getdate()) " & _
        ''response.Write sql
		consultar sql,RS
		
		
		SUMA_MontoSoles_LIMA_PART=0
		SUMA_MontoSoles_LIMA_PYME=0
		SUMA_Cantidad_LIMA_PART=0
		SUMA_Cantidad_LIMA_PYME=0
		SUMA_MontoSoles_LIMA_PART_CF=0
		SUMA_MontoSoles_LIMA_PYME_CF=0
		SUMA_Cantidad_LIMA_PART_CF=0
		SUMA_Cantidad_LIMA_PYME_CF=0
		
		SUMA_MontoSoles_PROV_PART=0
		SUMA_MontoSoles_PROV_PYME=0
		SUMA_Cantidad_PROV_PART=0
		SUMA_Cantidad_PROV_PYME=0
		SUMA_MontoSoles_PROV_PART_CF=0
		SUMA_MontoSoles_PROV_PYME_CF=0
		SUMA_Cantidad_PROV_PART_CF=0
		SUMA_Cantidad_PROV_PYME_CF=0
		
		SUMA_MontoSoles_PART=0
		SUMA_Cantidad_PART=0
		SUMA_MontoSoles_PYME=0
		SUMA_Cantidad_PYME=0
		SUMA_MontoSoles_PART_CF=0
		SUMA_Cantidad_PART_CF=0
		SUMA_MontoSoles_PYME_CF=0
		SUMA_Cantidad_PYME_CF=0
								
        RS.Filter=" Plaza='LIMA' "
        Do While not RS.EOF
		    SUMA_MontoSoles_LIMA_PART=SUMA_MontoSoles_LIMA_PART + RS.Fields("MontoSolesPART")
		    SUMA_Cantidad_LIMA_PART=SUMA_Cantidad_LIMA_PART + RS.Fields("CasosPART")
		    SUMA_MontoSoles_LIMA_PART_CF=SUMA_MontoSoles_LIMA_PART_CF + RS.Fields("MontoSolesConFasePART")
		    SUMA_Cantidad_LIMA_PART_CF=SUMA_Cantidad_LIMA_PART_CF + RS.Fields("CasosConFasePART")
		    SUMA_MontoSoles_LIMA_PYME=SUMA_MontoSoles_LIMA_PYME + RS.Fields("MontoSolesPYME")
		    SUMA_Cantidad_LIMA_PYME=SUMA_Cantidad_LIMA_PYME + RS.Fields("CasosPYME")
		    SUMA_MontoSoles_LIMA_PYME_CF=SUMA_MontoSoles_LIMA_PYME_CF + RS.Fields("MontoSolesConFasePYME")
		    SUMA_Cantidad_LIMA_PYME_CF=SUMA_Cantidad_LIMA_PYME_CF + RS.Fields("CasosConFasePYME")
        RS.MoveNeXt
        Loop 
        RS.Filter=""
        RS.Filter=" Plaza='PROVINCIA' "
        Do While not RS.EOF
		    SUMA_MontoSoles_PROV_PART=SUMA_MontoSoles_PROV_PART + RS.Fields("MontoSolesPART")
		    SUMA_Cantidad_PROV_PART=SUMA_Cantidad_PROV_PART + RS.Fields("CasosPART")
		    SUMA_MontoSoles_PROV_PART_CF=SUMA_MontoSoles_PROV_PART_CF + RS.Fields("MontoSolesConFasePART")
		    SUMA_Cantidad_PROV_PART_CF=SUMA_Cantidad_PROV_PART_CF + RS.Fields("CasosConFasePART")
		    SUMA_MontoSoles_PROV_PYME=SUMA_MontoSoles_PROV_PYME + RS.Fields("MontoSolesPYME")
		    SUMA_Cantidad_PROV_PYME=SUMA_Cantidad_PROV_PYME + RS.Fields("CasosPYME")
		    SUMA_MontoSoles_PROV_PYME_CF=SUMA_MontoSoles_PROV_PYME_CF + RS.Fields("MontoSolesConFasePYME")
		    SUMA_Cantidad_PROV_PYME_CF=SUMA_Cantidad_PROV_PYME_CF + RS.Fields("CasosConFasePYME")
        RS.MoveNeXt
        Loop    

		SUMA_MontoSoles_PART=SUMA_MontoSoles_LIMA_PART + SUMA_MontoSoles_PROV_PART
		SUMA_Cantidad_PART=SUMA_Cantidad_LIMA_PART + SUMA_Cantidad_PROV_PART
		SUMA_MontoSoles_PYME=SUMA_MontoSoles_LIMA_PYME + SUMA_MontoSoles_PROV_PYME
		SUMA_Cantidad_PYME=SUMA_Cantidad_LIMA_PYME + SUMA_Cantidad_PROV_PYME
		SUMA_MontoSoles_PART_CF=SUMA_MontoSoles_LIMA_PART_CF + SUMA_MontoSoles_PROV_PART_CF
		SUMA_Cantidad_PART_CF=SUMA_Cantidad_LIMA_PART_CF + SUMA_Cantidad_PROV_PART_CF
		SUMA_MontoSoles_PYME_CF=SUMA_MontoSoles_LIMA_PYME_CF + SUMA_MontoSoles_PROV_PYME_CF
		SUMA_Cantidad_PYME_CF=SUMA_Cantidad_LIMA_PYME_CF + SUMA_Cantidad_PROV_PYME_CF
		%>
		
		
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
			<title>SAE por Estudios</title>
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
						<font size=4 color=#483d8b face=Arial><b>Alertas por Estudio al <%=fechacierre%></b></font></td>
                        
					</tr>
					</table>
					<div id="cabtabla_rep" style="overflow:auto; height:auto; padding:0;"><!--margin-right: 17px;">-->
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#007DC5" rowspan=3>
							<font size=1 color=#FFFFFF face=Arial><b>ESTUDIOS</b></font>
						</td>
						<td bgcolor="#007DC5" colspan=7 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>PARTICULARES</b></font>
						</td>	
						<td bgcolor="#007DC5" colspan=7 align="center">
							<font size=1 color=#FFFFFF face=Arial><b>PYMES</b></font>
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
					</tr>
					<tr>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" width=15>
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>Miles S/.</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
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
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>Total</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b>100.00%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_CF*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center" width=15><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=&segmento=pyme&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PYME_CF*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PYME_CF*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>
					</tr>					
					<!--LIMA-->
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Lima</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PART%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PART_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PYME%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PYME_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=lima&segmento=pyme&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
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
		                        <a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=&confase="><font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESTUDIO")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPART")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePART")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=particulares&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPYME")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePYME")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=lima&segmento=pyme&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td><%end if%>
                        </tr>	                        
                        <%
                        RS.MoveNext
                        Loop 					
					%>			
					<!--PROVINCIA-->	
					<tr>
						<td bgcolor="#BEE8FB">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=&confase="><font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Provincia</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PART%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PART_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1),2)%>%</b></font></a>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=particulares&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PYME%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</b></font></a>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME_CF/1000,0)%></b></font></a>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PYME_CF%></b></font></a>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1),2)%>%</b></font></a>
						</td>
						<td bgcolor="#BEE8FB" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=&plaza=provincia&segmento=pyme&confase=1"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=90,"[90%-100%]",iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a></td>
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
		                        <a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=&confase="><font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESTUDIO")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPART")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePART")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1),2)%>%</font></a><%end if%>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="center"><%if RS.Fields("CasosPART")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=particulares&confase=1"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0></a><%end if%></td>
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPYME")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</font></a><%end if%>
	                        </td>		
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")/1000,0)%></font></a><%end if%>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePYME")%></font></a><%end if%>
	                        </td>				
	                        <td bgcolor="<%=bgcolor%>" align="right">
		                        <%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase=1"><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1),2)%>%</font></a><%end if%>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="center"><a href="repsaecasosdet.asp?paginapadre=repsaeestudio.asp&asignado=S&fechacierre=<%=fechacierre%>&codestudio=<%=trim(RS.Fields("CODDATO"))%>&plaza=provincia&segmento=pyme&confase=1"><%if RS.Fields("CasosPYME")=0 then%>&nbsp;<%else%><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" alt="<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"[90%-100%]",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"[80%-90%]","[0%-80%&gt;"))%>" width="12" border=0><%end if%></a></td>
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

