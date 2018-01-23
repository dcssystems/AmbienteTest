<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repidasegmentobec.asp") then
		
	    fechacierre=obtener("fechacierre")
	    ''if not IsDate(fechacierre) then
	    if fechacierre="" then
            sql="select max(fechadatos) as fechadatos from FB.CentroCobranzas.dbo.pd_casos_jud"
            consultar sql,RS	
            fechacierre=RS.Fields("fechadatos")
            RS.Close
		end if
		
		sql="select MAX(fechacierre) as FechaCierre from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'"
        consultar sql,RS	
        fechacierremesant=RS.Fields("FechaCierre")
        RS.Close		
		%>
		
		
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		    <!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
			<title>IDA Mora</title>
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
            var columnas=20; //CANTIDAD DE COLUMNAS//

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
						<font size=4 color=#483d8b face=Arial><b>Informe Diario de Mora (IDA - BC) al <%=fechacierre%> - BEC</b></font></td>
					</tr>
					</table>
					<div id="cabtabla_rep" style="overflow:auto; height:auto; padding:0;"><!--margin-right: 17px;">-->
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#007DC5">
							<font size=1 color=#FFFFFF face=Arial><b>Especialista</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Stock Cierre</b></font>
						</td>	
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Ingresos Mes</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Meta Mensual</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Meta Anual**</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Avance Anual con Venta</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Avance Anual sin Venta</b></font>
						</td>												
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Avance Anual s/ Venta %</b></font>
						</td>																
						<!--<td bgcolor="#007DC5" align="center" width="15">
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>-->																
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Avance Mes con Venta</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Avance Mes sin Venta</b></font>
						</td>												
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Avance Mes s/ Venta %</b></font>
						</td>																
						<td bgcolor="#007DC5" align="center" width="15">
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>
						<td bgcolor="#FFFFFF" align="center" width="1%">
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>	
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>REF</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>TRA</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>EJEC</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>EFE</b></font>
						</td>																											
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>OTRO</b></font>
						</td>							
						<td bgcolor="#FFFFFF" align="center" width="1%">
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>	
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>BEC</b></font>
						</td>												
					</tr>		
					</table>
					</div>
					<div id="dettabla_rep" style="overflow:auto; height:80%; padding:0;">
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<%
					
		                sql="If OBJECT_ID('TempDB.dbo.#INGRESOS') IS NOT NULL DROP TABLE #INGRESOS"
	                    conn.execute sql
	                    
		                sql="If OBJECT_ID('TempDB.dbo.#CasosJudDia') IS NOT NULL DROP TABLE #CasosJudDia"
	                    conn.execute sql
	                    
	                    sql="select * into #CasosJudDia from FB.CentroCobranzas.dbo.PD_Casos_Jud where SEGMENTO_RIESGO in ('BEC','BEC-RED','PROMOTORES') and FECHADATOS='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'"
	                    conn.execute sql

		                sql="If OBJECT_ID('TempDB.dbo.#SALIDAS') IS NOT NULL DROP TABLE #SALIDAS"
	                    conn.execute sql	                    
	                    
		                ''sql="If OBJECT_ID('TempDB.dbo.#SAL_AAAA') IS NOT NULL DROP TABLE #SAL_AAAA"
	                    ''conn.execute sql		                    
	                    
					    sql="SELECT " &_
			                    "    A.cod_territorio," &_
			                    "    A.TERRITORIO_HOMOLOGADO," &_
			                    "    D.ESPECIALISTA_JUD_BEC," &_
			                    "    ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO," &_
			                    "    SUM(A.soles + A.dolares*A.tipocambio) as Ingresos " &_
		                        "into #INGRESOS " &_
		                        "FROM FB.CentroCobranzas.dbo.Vista_Ingresos_Segmento A " &_
			                    "    LEFT JOIN	#CasosJudDia C	ON C.CODCENTRAL	= A.CODCENTRAL  " &_
								"					                        AND C.CONTRATO		= A.CONT_SAE " &_
								"LEFT JOIN FB.CentroCobranzas.dbo.PD_CENTROS_ESPECIALISTA D ON A.CODOFICINA = D.COD_OFICINA " &_
		                        "WHERE	A.FECHA_INGRESO>(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "')" &_
				                "        AND A.FECHA_INGRESO<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
				                "        AND ltrim(rtrim(A.segmento_riesgo)) in ('BEC','BEC-RED','PROMOTORES') " &_
				                "        AND ltrim(rtrim(A.ENTIDAD))			in ('BC','35') " &_
		                        "GROUP BY	A.cod_territorio, " &_
					            "            A.TERRITORIO_HOMOLOGADO, " &_
			                    "    D.ESPECIALISTA_JUD_BEC, ltrim(rtrim(A.SEGMENTO_RIESGO)) "

					    conn.execute sql
					    

		                sql="SELECT	 " &_
			                        "B.cod_territorio, " &_
			                        "B.TERRITORIO_HOMOLOGADO, " &_
			                        "SUM(soles + dolares*tipocambio) as Salidas, " &_
			                        "SUM(CASE WHEN MODALIDAD						=	'R'						THEN soles + dolares*tipocambio ELSE 0 END) as REF, " &_
			                        "SUM(CASE WHEN MODALIDAD						=	'T'						THEN soles + dolares*tipocambio ELSE 0 END) as TRA, " &_
			                        "SUM(CASE WHEN MODALIDAD						IN	('N','M','G')			THEN soles + dolares*tipocambio ELSE 0 END) as EJEC, " &_
			                        "SUM(CASE WHEN MODALIDAD						=	'D'						THEN soles + dolares*tipocambio ELSE 0 END) as EFE, " &_
			                        "SUM(CASE WHEN MODALIDAD						IN	('B','S','E','Y','Z')	THEN soles + dolares*tipocambio ELSE 0 END)	as OTRO, " &_
			                        "ltrim(rtrim(B.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
		                        "INTO #SALIDAS " &_
		                        "FROM FB.CentroCobranzas.dbo.XCOM_MAE_SALIDAS A " &_
			                    "    LEFT JOIN	FB.CentroCobranzas.dbo.MAE_TerritorioSegmento  B	on A.CODCENTRAL		= B.codcentral  " &_
								"					                        AND A.CONT_SAE		= B.contrato " &_
			                    "    LEFT JOIN	#CasosJudDia C	ON C.CODCENTRAL	= A.CODCENTRAL  " &_
								"					                        AND C.CONTRATO		= A.CONT_SAE " &_
		                        "WHERE	A.FECHA_SALIDAS>(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') " &_
				                "        AND A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
				                "        AND RTRIM(LTRIM(B.SEGMENTO_RIESGO))	in ('BEC','BEC-RED','PROMOTORES')  " &_
				                "        AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
				                "        AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z') " &_
		                        "GROUP BY B.cod_territorio, " &_
				                "         B.TERRITORIO_HOMOLOGADO, "	&_
				                "         ltrim(rtrim(B.SEGMENTO_RIESGO))"
					
					    '''conn.execute sql					


		                sql="SELECT	top 0 " &_
			                "        B.cod_territorio,	" &_
			                "        B.TERRITORIO_HOMOLOGADO,	" &_
			                "        SUM(soles + dolares*tipocambio) as Salidas,	" &_
			                "        ltrim(rtrim(B.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
		                    "    INTO #SAL_AAAA " &_
		                    "    FROM FB.CentroCobranzas.dbo.XCOM_MAE_SALIDAS A	" &_
			                "        LEFT JOIN	FB.CentroCobranzas.dbo.MAE_TerritorioSegmento  B	on A.CODCENTRAL		= B.codcentral 	" &_
						    "						                        AND A.CONT_SAE		= B.contrato	" &_
			                "        LEFT JOIN	#CasosJudDia C	ON C.CODCENTRAL	= A.CODCENTRAL	" &_ 
							"						                        AND C.CONTRATO		= A.CONT_SAE	" &_
		                    "    WHERE	YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "')	" &_
				            "            AND A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'	" &_
				            "            AND RTRIM(LTRIM(B.SEGMENTO_RIESGO))	in ('BEC','BEC-RED','PROMOTORES') 	" &_
				            "            AND ltrim(rtrim(A.entidad)) in ('BC','35')		" &_
				            "            AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z')	" &_
		                    "    GROUP BY B.cod_territorio,	" &_
				            "             B.TERRITORIO_HOMOLOGADO,	" &_
				            "             ltrim(rtrim(B.SEGMENTO_RIESGO))"  
					    ''response.Write now() & "<BR>"
					    '''conn.execute sql					
					    ''response.Write now() & "<BR>"
					    
		                sql="SELECT	" &_
			                "B.cod_territorio, " &_
			                "B.TERRITORIO_HOMOLOGADO, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' AND MODALIDAD                     NOT IN ('W')                 THEN soles + dolares*tipocambio ELSE 0 END) as Salidas, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' THEN soles + dolares*tipocambio ELSE 0 END) as SalidasConVenta, " &_
			                "SUM(CASE WHEN  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END)           as SalidasAAAA, " &_
			                "SUM(soles + dolares*tipocambio) as SalidasAAAAConVenta, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						=	'R'						THEN soles + dolares*tipocambio ELSE 0 END) as REF, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						=	'T'						THEN soles + dolares*tipocambio ELSE 0 END) as TRA, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						IN	('N','M','G')			THEN soles + dolares*tipocambio ELSE 0 END) as EJEC, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						=	'D'						THEN soles + dolares*tipocambio ELSE 0 END) as EFE, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						IN	('B','S','E','Y','Z')	THEN soles + dolares*tipocambio ELSE 0 END)	as OTRO, " &_
			                "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
		                "INTO #SALIDAS " &_
		                "FROM FB.CentroCobranzas.dbo.XCOM_MAE_SALIDAS A " &_
			                "LEFT OUTER JOIN	FB.CentroCobranzas.dbo.MAE_TerritorioSegmento  B	on A.CODCENTRAL		= B.codcentral  " &_
							"						                AND A.CONT_SAE		= B.contrato " &_
			                "LEFT OUTER JOIN #CasosJudDia C	ON C.CODCENTRAL	= A.CODCENTRAL  " &_
							"						                AND C.CONTRATO		= A.CONT_SAE " &_
		                "WHERE	YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') " &_
				        "        AND A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
				        "        AND RTRIM(LTRIM(B.SEGMENTO_RIESGO))	in ('BEC','BEC-RED','PROMOTORES')  " &_
				        "        AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
				        "        AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z') " &_
		                "GROUP BY B.cod_territorio, " &_
				        "         B.TERRITORIO_HOMOLOGADO, " &_
				        "         ltrim(rtrim(A.SEGMENTO_RIESGO))"
				        
				        
				        
                        sql="SELECT	" &_
			                "A.cod_territorio, " &_
			                "A.TERRITORIO_HOMOLOGADO, " &_
			                "    D.ESPECIALISTA_JUD_BEC," &_			                
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' AND MODALIDAD                     NOT IN ('W')                 THEN soles + dolares*tipocambio ELSE 0 END) as Salidas, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' THEN soles + dolares*tipocambio ELSE 0 END) as SalidasConVenta, " &_
			                "SUM(CASE WHEN  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END)           as SalidasAAAA, " &_
			                "SUM(soles + dolares*tipocambio) as SalidasAAAAConVenta, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						=	'R'						THEN soles + dolares*tipocambio ELSE 0 END) as REF, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						=	'T'						THEN soles + dolares*tipocambio ELSE 0 END) as TRA, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						IN	('N','M','G')			THEN soles + dolares*tipocambio ELSE 0 END) as EJEC, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						=	'D'						THEN soles + dolares*tipocambio ELSE 0 END) as EFE, " &_
			                "SUM(CASE WHEN A.FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' and MODALIDAD						IN	('B','S','E','Y','Z')	THEN soles + dolares*tipocambio ELSE 0 END)	as OTRO, " &_
			                "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
		                "INTO #SALIDAS " &_
		                "FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A " &_
			                "LEFT JOIN #CasosJudDia C	ON C.CODCENTRAL	= A.CODCENTRAL  " &_
							"						                AND C.CONTRATO		= A.CONT_SAE " &_
							"LEFT JOIN FB.CentroCobranzas.dbo.PD_CENTROS_ESPECIALISTA D ON A.CODOFICINA = D.COD_OFICINA " &_
		                "WHERE	YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') " &_
				        "        AND A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
				        "        AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('BEC','BEC-RED','PROMOTORES')  " &_
				        "        AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
				        "        AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z','W') " &_
		                "GROUP BY A.cod_territorio, " &_
				        "         A.TERRITORIO_HOMOLOGADO, " &_
				        "         ltrim(rtrim(A.SEGMENTO_RIESGO))," &_
		                "    D.ESPECIALISTA_JUD_BEC"			        				        
				        
				        conn.execute sql		
				        
					    
                        sql="select codterritorio,territorio,especialista,SUM(jud) as StockCierre,IsNull((select sum(ingresos) from #INGRESOS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as INGRESOS,IsNull((select sum(salidas) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as SALIDAS,IsNull((select sum(salidasconventa) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as SALIDASCONVENTA,IsNull((select sum(REF) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as REF,IsNull((select sum(TRA) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as TRA,IsNull((select sum(EJEC) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as EJEC,IsNull((select sum(EFE) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as EFE,IsNull((select sum(OTRO) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as OTRO,0 as PART,IsNull((select sum(Salidas) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as BEC,IsNull((select sum(SalidasAAAA) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as SalidasAAAA,IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where cod_territorio=A.codterritorio and SEGMENTO_RIESGO=A.segmento_riesgo and ESPECIALISTA_JUD_BEC=A.Especialista),0) as SalidasAAAAConVenta, " &_
                            " IsNull((select sum(meta) from FB.CentroCobranzas.dbo.Metas_IDA where codterritorio=A.codterritorio and segmento=A.segmento_riesgo),0) as MetaMes  " & _
                            "from FB.CentroCobranzas.dbo.PD_Casos_Jud A where segmento_riesgo='BEC' and fechadatos=(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') group by codterritorio,territorio,especialista,segmento_riesgo order by StockCierre desc"
                        ''response.Write sql
                        bgcolor="#FFFFFF"
                        consultar sql,RS	
                        
                        SUMA_STOCK=0
                        SUMA_INGMES=0
                        SUMA_METAMES=0
                        SUMA_METAAAAA=0
                        SUMA_AVANCEAAAA=0
                        SUMA_AVANCEAAAACONVENTA=0
                        SUMA_AVANCEAAAAPORC=0
                        SUMA_AVANCEAAAAPORCCONVENTA=0
                        SUMA_AVANCEMES=0
                        SUMA_AVANCEMESCONVENTA=0
                        SUMA_AVANCEMESPORC=0
                        SUMA_AVANCEMESPORCCONVENTA=0
                        SUMA_REF=0
                        SUMA_TRA=0
                        SUMA_EJEC=0
                        SUMA_EFE=0
                        SUMA_OTRO=0
                        SUMA_BEC=0
                        
                        Do While not RS.EOF
                            if bgcolor="#FFFFFF"then
                                bgcolor="#F5F5F5"
                            else
                                bgcolor="#FFFFFF"
                            end if
                        %>
                        <tr>
	                        <td bgcolor="<%=bgcolor%>">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=RS.Fields("Especialista") & " (" & RS.Fields("TERRITORIO") & ")"%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("StockCierre")/1000,0)%></font></a>
	                        </td>						
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("Ingresos")/1000,0)%></font></a>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MetaMes")/1000,0)%></font></a>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MetaMes")*12/1000,0)%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasAAAAConVenta")/1000,0)%></font></a>
	                        </td>		                        
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasAAAA")/1000,0)%></font></a>
	                        </td>	                        						                        	                        
						    <td bgcolor="<%=bgcolor%>" align="right" width="5%">
							    <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0),2)%>%</font></a>
						    </td>	
						    <!--<td bgcolor="<%=bgcolor%>" align="center" width="15"><a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><img src="imagenes/<%=iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)>=90,"sem_verde",iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)>=90,"[90%-100%]",iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)>=90,"[90%-100%]",iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>-->
						    <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasConVenta")/1000,0)%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("Salidas")/1000,0)%></font></a>
	                        </td>	                        						                        
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
							    <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codestudio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0),2)%>%</font></a>
						    </td>	
						    <td bgcolor="<%=bgcolor%>" align="center" width="15"><a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><img src="imagenes/<%=iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)>=90,"sem_verde",iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)>=90,"[90%-100%]",iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)>=90,"[90%-100%]",iif(iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0></a></td>	                        
						    <td bgcolor="#FFFFFF" align="center" width="1%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>						    
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("REF")/1000,0)%></font></a>
	                        </td>	 
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("TRA")/1000,0)%></font></a>
	                        </td>	 
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("EJEC")/1000,0)%></font></a>
	                        </td>	 
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("EFE")/1000,0)%></font></a>
	                        </td>	 	                        	                        	                        	                        
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("OTRO")/1000,0)%></font></a>
	                        </td>	 	                        	 
						    <td bgcolor="#FFFFFF" align="center" width="1%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	                                               	                        	                        	                        
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <a href="repsaecasosempdet.asp?paginapadre=repidasegmentobec.asp&asignado=&fechacierre=<%=fechacierre%>&codterritorio=<%=RS.Fields("CODTERRITORIO")%>&plaza=&segmento=BEC&confase="><font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("BEC")/1000,0)%></font></a>
	                        </td>		                        	                        
                        </tr>	                        
                        <%

                        SUMA_STOCK=SUMA_STOCK + RS.Fields("StockCierre")/1000
                        SUMA_INGMES=SUMA_INGMES + RS.Fields("Ingresos")/1000
                        SUMA_METAMES=SUMA_METAMES + RS.Fields("MetaMes")/1000
                        SUMA_METAAAAA=SUMA_METAAAAA + RS.Fields("MetaMes")*12/1000
                        SUMA_AVANCEAAAACONVENTA=SUMA_AVANCEAAAACONVENTA + RS.Fields("SalidasAAAAConVenta")/1000
                        SUMA_AVANCEAAAA=SUMA_AVANCEAAAA + RS.Fields("SalidasAAAA")/1000
                        SUMA_AVANCEAAAAPORC=SUMA_AVANCEAAAAPORC + iif(RS.Fields("MetaMes")>0,iif(RS.Fields("SalidasAAAA")>0,RS.Fields("SalidasAAAA")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes")*12,1),0),0)
                        SUMA_AVANCEMESCONVENTA=SUMA_AVANCEMESCONVENTA + RS.Fields("SalidasConVenta")/1000
                        SUMA_AVANCEMES=SUMA_AVANCEMES + RS.Fields("Salidas")/1000
                        SUMA_AVANCEMESPORC=SUMA_AVANCEMESPORC + iif(RS.Fields("MetaMes")>0,iif(RS.Fields("Salidas")>0,RS.Fields("Salidas")*100/iif(RS.Fields("MetaMes")>0,RS.Fields("MetaMes"),1),0),0)
                        SUMA_REF=SUMA_REF + RS.Fields("REF")/1000
                        SUMA_TRA=SUMA_TRA + RS.Fields("TRA")/1000
                        SUMA_EJEC=SUMA_EJEC + RS.Fields("EJEC")/1000
                        SUMA_EFE=SUMA_EFE + RS.Fields("EFE")/1000
                        SUMA_OTRO=SUMA_OTRO + RS.Fields("OTRO")/1000
                        SUMA_BEC=SUMA_BEC + RS.Fields("BEC")/1000
                        
                        RS.MoveNext
                        Loop
                        RS.Close 					
					    %>			
                        <tr>
						    <td bgcolor="#007DC5">
							    <font size=1 color=#FFFFFF face=Arial><b>TOTAL</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_STOCK,0)%></b></font>
						    </td>	
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_INGMES,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_METAMES,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_METAAAAA,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_AVANCEAAAACONVENTA,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_AVANCEAAAA,0)%></b></font>
						    </td>												
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0),2)%>%</b></font>
						    </td>																
						    <!--<td bgcolor="#007DC5" align="center" width="15">
							    <img src="imagenes/<%=iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0)>=90,"sem_verde",iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0)>=90,"[90%-100%]",iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0)>=90,"[90%-100%]",iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEAAAA>0,SUMA_AVANCEAAAA*100/iif(SUMA_METAMES>0,SUMA_METAMES*12,1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0>
						    </td>-->
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_AVANCEMESCONVENTA,0)%></b></font>
						    </td>	
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_AVANCEMES,0)%></b></font>
						    </td>												
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0),2)%>%</b></font>
						    </td>																
						    <td bgcolor="#FFFFFF" align="center" width="15">
							    <img src="imagenes/<%=iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0)>=90,"sem_verde",iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0)>=80,"sem_amarillo","sem_rojo"))%>.png" title="<%=iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0)>=90,"[90%-100%]",iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" alt="<%=iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0)>=90,"[90%-100%]",iif(iif(SUMA_METAMES>0,iif(SUMA_AVANCEMES>0,SUMA_AVANCEMES*100/iif(SUMA_METAMES>0,SUMA_METAMES,1),0),0)>=80,"[80%-90%]","[0%-80&gt;"))%>" width="12" border=0>
						    </td>																
						    <td bgcolor="#FFFFFF" align="right" width="1%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_REF,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_TRA,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_EJEC,0)%></b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_EFE,0)%></b></font>
						    </td>																											
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_OTRO,0)%></b></font>
						    </td>							
						    <td bgcolor="#FFFFFF" align="right" width="1%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b><%=FormatNumber(SUMA_BEC,0)%></b></font>
						    </td>												
					    </tr>					    
					</table>
					<font size=1 color=#483d8b face=Arial><b>&nbsp;* Valores en Miles de Soles (SETC)</b></font>
					<BR>
					<font size=1 color=#483d8b face=Arial><b>&nbsp;** Meta mensual multiplicado por el nro de mes actual</b></font>
					</div>
					<input type="hidden" name="expimp" value="">		
		        <input type="hidden" name="pag" value="<%=pag%>">
		        <input type="hidden" name="fechacierre" value="<%=obtener("fechacierre")%>">
			</form>							
			</body>
		</html>	
		
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display = "none";</script>	
		<%	
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

