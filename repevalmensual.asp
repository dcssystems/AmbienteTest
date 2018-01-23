<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repevalmensual.asp") then
		
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

        

        sql="select top 5 * " &_
			 "from FB.CentroCobranzas.dbo.PD_FechasCierre  " &_
			 "where fechacierre<=  " &_                
			 "(select MAX(fechadatos) from FB.CentroCobranzas.dbo.PD_Casos_JUD)  " &_
			 "and month(fechacierre) in (12,3,6,9) " &_
			 "and fechacierre in (select fechadatos from FB.CentroCobranzas.dbo.PD_Casos_JUD) " &_
			 "order by fechacierre desc " 
		consultar sql,RS

        select case int(mid(fechacierre,4,2))
            case 1:
                    mes4="Diciembre"
                    mes3="Setiembre"
                    mes2="Junio"
                    mes1="Marzo"
            case 2:
                    mes4="Diciembre"
                    mes3="Setiembre"
                    mes2="Junio"
                    mes1="Marzo"
            case 3:
                    mes4="Marzo"
                    mes3="Diciembre"
                    mes2="Setiembre"
                    mes1="Junio"
                    
            case 4:
                    mes4="Marzo"
                    mes3="Diciembre"
                    mes2="Setiembre"
                    mes1="Junio"
            case 5:
                    mes4="Marzo"
                    mes3="Diciembre"
                    mes2="Setiembre"
                    mes1="Junio"
            case 6:
                    mes4="Junio"
                    mes3="Marzo"
                    mes2="Diciembre"
                    mes1="Setiembre"
                    
            case 7:
                    mes4="Junio"
                    mes3="Marzo"
                    mes2="Diciembre"
                    mes1="Setiembre"
            case 8:
                    mes4="Junio"
                    mes3="Marzo"
                    mes2="Diciembre"
                    mes1="Setiembre"
            case 9:
                    mes4="Setiembre"
                    mes3="Junio"
                    mes2="Marzo"
                    mes1="Diciembre"
            case 10:
                    mes4="Setiembre"
                    mes3="Junio"
                    mes2="Marzo"
                    mes1="Diciembre"
            case 11:
                    mes4="Setiembre"
                    mes3="Junio"
                    mes2="Marzo"
                    mes1="Diciembre"
            case 12:
                    mes4="Diciembre"
                    mes3="Setiembre"
                    mes2="Junio"
                    mes1="Marzo"
                    
        end select 
        
		counter=1
		f1=""
		f2=""
		f3=""
		f4=""
		f5=""

        Do While not RS.EOF
        	Select Case counter
        		case 1
        			f5=RS.Fields("FechaCierre")
        		case 2
        			f4=RS.Fields("FechaCierre")
        		case 3
        			f3=RS.Fields("FechaCierre")
        		case 4
        			f2=RS.Fields("FechaCierre")
        		case 5
        			f1=RS.Fields("FechaCierre")
        	End Select
        	counter=counter+1
			RS.MoveNext
		Loop
		RS.Close
		
		plaza=obtener("plaza")

		if plaza="" then
			plaza="LIMA"
		end if



		%>
		
		
		<!--Ojo esta ventana siempre es flotante-->
		<html>
		    <!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
			<title>Informe de Evaluación Mensual por Especialista</title>
			<script language=javascript>
			    var ventanaverestudio;
			    function inicio() {
			        dibujarTabla(0);
			    }
			    function modificar(xxestudio) {
			        ventanaverestudio = window.open("verestudio.asp?vistapadre=" + window.name + "&paginapadre=repevalmensual.asp&estudio=" + xxestudio + "&fechacierre=<%=fechacierre%>", "VerEstudio", "scrollbars=yes,scrolling=yes,top=" + ((screen.height) / 2 - 300) + ",height=600,width=" + (screen.width / 2 + 300) + ",left=" + (screen.width / 2 - 475) + ",resizable=yes");
			        ventanaverestudio.focus();
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
                var ancho1, ancho2, i;
                var columnas = 15; //CANTIDAD DE COLUMNAS//

                function ajustaCeldas() {
                    for (i = 0; i < columnas; i++) {
                        ancho1 = document.getElementById("encabezado").rows.item(0).cells.item(i).offsetWidth;
                        ancho2 = document.getElementById("datos").rows.item(0).cells.item(i).offsetWidth;
                        if (ancho1 > ancho2) {
                            document.getElementById("datos").rows.item(0).cells.item(i).width = ancho1 - 6;
                        }
                        else {
                            document.getElementById("encabezado").rows.item(0).cells.item(i).width = ancho2 - 6;
                        }
                    }
                }

                function cuadratabla() {
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
					    <td bgcolor="#F5F5F5" width=90>&nbsp;<input name="fechacierrebusc" id="fechacierrebusc" readonly type=text maxlength=10 size=10 value="<%=fechacierre%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('fechacierrebusc', '%d/%m/%Y');"></td>
						<td bgcolor="#F5F5F5" width=160>
	                        <select name="plaza" style="font-size: x-small; width: 120px;">
			                    <option value="LIMA" <%if plaza="LIMA" then%> selected<%end if%>>LIMA</option>
			                    <option value="PROVINCIA" <%if plaza="PROVINCIA" then%> selected<%end if%>>PROVINCIA</option>
		                    </select>&nbsp;&nbsp;<a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle height="18"></a>
	                    </td>		
						<td bgcolor="#F5F5F5" align=center><font size=4 color=#483d8b face=Arial><b>Informe de Evaluación Mensual - <%=plaza%></b></font></td>
					</tr>
					</table>
					<div id="cabtabla_rep" style="overflow:auto; height:auto; padding:0;"><!--margin-right: 17px;">-->
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<tr>
						<td bgcolor="#007DC5">
							<font size=1 color=#FFFFFF face=Arial><b>Estudio</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="10%">
							<font size=1 color=#FFFFFF face=Arial><b>Segmento</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>N° Clientes</b></font>
						</td>	
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Total Provisión</b></font>
						</td>	
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Total Asignado</b></font>
						</td>
						<td bgcolor="#FFFFFF" align="center" width="1%">
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Total Salidas Año Sin Venta</b></font>
						</td>													
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Total Salidas Año Con Venta</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Total Salidas Mes Sin Venta</b></font>
						</td>													
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Total Salidas Mes Con Venta</b></font>
						</td>	
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>Meta Mensual</b></font>
						</td>	
						<%if f1<>"" then %>
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes1%></b></font>
							</td>
							<td bgcolor="#FFFFFF" align="center" width="2%">
								<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							</td>
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes2%></b></font>
							</td>
							<td bgcolor="#FFFFFF" align="center" width="2%">
								<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							</td>	
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes3%></b></font>
							</td>
							<td bgcolor="#FFFFFF" align="center" width="2%">
								<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							</td>
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes4%></b></font>
							</td>
							<td bgcolor="#FFFFFF" align="center" width="2%">
								<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							</td>
						<%else %>
							<%if f2<>"" then %>
								<td bgcolor="#007DC5" align="center" width="5%">
									<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes2%></b></font>
								</td>
								<td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>	
								<td bgcolor="#007DC5" align="center" width="5%">
									<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes3%></b></font>
								</td>
								<td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>
								<td bgcolor="#007DC5" align="center" width="5%">
									<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes4%></b></font>
								</td>
								<td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>
							<%else %>
								<%if f3<>"" then %>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes3%></b></font>
									</td>
									<td bgcolor="#FFFFFF" align="center" width="2%">
										<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
									</td>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes4%></b></font>
									</td>
									<td bgcolor="#FFFFFF" align="center" width="2%">
										<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
									</td>
								<%else %>
									<%if f4<>"" then %>
										<td bgcolor="#007DC5" align="center" width="5%">
											<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta <%=mes4%></b></font>
										</td>
										<td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							            </td>
									<%end if %>

								<%end if %>

							<%end if %>

						<%end if %>									
						
						<td bgcolor="#007DC5" align="center" width="5%">
							<font size=1 color=#FFFFFF face=Arial><b>% Avance Anual</b></font>
						</td>

						<%if f1<>"" then %>
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes1%></b></font>
							</td>
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes2%></b></font>
							</td>	
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes3%></b></font>
							</td>
							<td bgcolor="#007DC5" align="center" width="5%">
								<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes4%></b></font>
							</td>
						<%else %>
							<%if f2<>"" then %>
								<td bgcolor="#007DC5" align="center" width="5%">
									<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes2%></b></font>
								</td>	
								<td bgcolor="#007DC5" align="center" width="5%">
									<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes3%></b></font>
								</td>
								<td bgcolor="#007DC5" align="center" width="5%">
									<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes4%></b></font>
								</td>
							<%else %>
								<%if f3<>"" then %>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes3%></b></font>
									</td>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes4%></b></font>
									</td>
								<%else %>
									<%if f4<>"" then %>
										<td bgcolor="#007DC5" align="center" width="5%">
											<font size=1 color=#FFFFFF face=Arial><b>% Efectividad Estudio <%=mes4%></b></font>
										</td>
									<%end if %>

								<%end if %>

							<%end if %>

						<%end if %>
											
																							
					</tr>		
					</table>
					</div>
					<div id="dettabla_rep" style="overflow:auto; height:80%; padding:0;">
					<table border=0 cellspacing=2 cellpadding=4 width="100%">
					<%
					
		                sql="If OBJECT_ID('TempDB.dbo.#INGRESOS') IS NOT NULL DROP TABLE #INGRESOS"
	                    conn.execute sql
	                    
		                ''sql="If OBJECT_ID('TempDB.dbo.#CasosJudDia') IS NOT NULL DROP TABLE #CasosJudDia"
	                    ''conn.execute sql
	                    
	                    ''sql="select * into #CasosJudDia from FB.CentroCobranzas.dbo.PD_Casos_Jud where FECHADATOS='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "'"
	                    ''response.Write sql
	                    ''conn.execute sql

		                sql="If OBJECT_ID('TempDB.dbo.#SALIDAS') IS NOT NULL DROP TABLE #SALIDAS"
	                    conn.execute sql	

	                    sql="If OBJECT_ID('TempDB.dbo.#STOCKS') IS NOT NULL DROP TABLE #STOCKS"
	                    conn.execute sql                   
	                    
		                ''sql="If OBJECT_ID('TempDB.dbo.#SAL_AAAA') IS NOT NULL DROP TABLE #SAL_AAAA"
	                    ''conn.execute sql		                    
	                    
					   	if f1<>"" then
					    
					     sql="SELECT " &_
			                    "    IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio," &_
			                    "    SUM(CASE WHEN A.FECHA_INGRESO> '" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "'  AND A.FECHA_INGRESO<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' THEN A.soles + A.dolares*A.tipocambio ELSE 0 END) as Ingresos, " &_
			                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT4, " &_
			                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT3, " &_
			                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT2, " &_
			                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f1,7,4) & mid(f1,4,2) & mid(f1,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT1, " &_
			                    "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
		                        "into #INGRESOS " &_
		                        "FROM FB.CentroCobranzas.dbo.Vista_Ingresos_Segmento A " &_
			                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE  " &_
			                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA " &_
		                        "WHERE ltrim(rtrim(A.segmento_riesgo)) in ('PARTICULARES', 'PYME') " &_
				                "        AND ltrim(rtrim(A.ENTIDAD))			in ('BC','35') " &_
		                        "GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
								"ltrim(rtrim(A.SEGMENTO_RIESGO)) "
                        ''response.Write sql
					    conn.execute sql

						    
						else
							if f2 <>"" then
							    sql="SELECT " &_
					                    "    IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio," &_
					                    "    SUM(CASE WHEN A.FECHA_INGRESO> '" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "'  AND A.FECHA_INGRESO<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' THEN A.soles + A.dolares*A.tipocambio ELSE 0 END) as Ingresos, " &_
					                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT4, " &_
			                            "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT3, " &_
			                            "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT2, " &_
			                            "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
				                        "into #INGRESOS " &_
				                        "FROM FB.CentroCobranzas.dbo.Vista_Ingresos_Segmento A " &_
					                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE  " &_
					                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA " &_
				                        "WHERE ltrim(rtrim(A.segmento_riesgo)) in ('PARTICULARES', 'PYME') " &_
						                "        AND ltrim(rtrim(A.ENTIDAD))			in ('BC','35') " &_
				                        "GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
										"ltrim(rtrim(A.SEGMENTO_RIESGO)) "
		                        ''response.Write sql
							    conn.execute sql

							else
								if f3<>"" then
										sql="SELECT " &_
				                    "    IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio," &_
				                    "    SUM(CASE WHEN A.FECHA_INGRESO> '" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "'  AND A.FECHA_INGRESO<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' THEN A.soles + A.dolares*A.tipocambio ELSE 0 END) as Ingresos, " &_
				                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT4, " &_
		                            "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT3, " &_
		                            "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
			                        "into #INGRESOS " &_
			                        "FROM FB.CentroCobranzas.dbo.Vista_Ingresos_Segmento A " &_
				                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE  " &_
				                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA " &_
			                        "WHERE ltrim(rtrim(A.segmento_riesgo)) in ('PARTICULARES', 'PYME') " &_
					                "        AND ltrim(rtrim(A.ENTIDAD))			in ('BC','35') " &_
			                        "GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
									"ltrim(rtrim(A.SEGMENTO_RIESGO)) "
			                        ''response.Write sql
								    conn.execute sql

								else

									if f4<>"" then
						    			
							    		sql="SELECT " &_
					                    "    IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio," &_
					                    "	SUM(CASE WHEN A.FECHA_INGRESO > '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' and A.FECHA_INGRESO <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT4, " &_
		                                "	SUM(CASE WHEN A.FECHA_INGRESO between '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' and '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "'  THEN A.soles + A.dolares*A.tipocambio ELSE 0  END) as IT4, " &_
										"ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
				                        "into #INGRESOS " &_
				                        "FROM FB.CentroCobranzas.dbo.Vista_Ingresos_Segmento A " &_
					                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO		= A.CONT_SAE  " &_
					                    "    LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO		= C.SIGLA " &_
				                        "WHERE ltrim(rtrim(A.segmento_riesgo)) in ('PARTICULARES', 'PYME') " &_
						                "        AND ltrim(rtrim(A.ENTIDAD))			in ('BC','35') " &_
				                        "GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
										"ltrim(rtrim(A.SEGMENTO_RIESGO)) "
				                        ''response.Write sql
									    conn.execute sql
									end if
								end if
							end if
						end if
				        
			

						if f1<>"" then
					        sql= "SELECT " &_
								"IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio, " &_
								"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') and  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAA, " &_ 
								"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAAConVenta,   " &_ 
								"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as Salidas, " &_
								"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' THEN soles + dolares*tipocambio ELSE 0 END) as SalidaConVenta, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  FECHA_SALIDAS <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S4T, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' AND FECHA_SALIDAS <= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S3T, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' AND FECHA_SALIDAS <= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S2T, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f1,7,4) & mid(f1,4,2) & mid(f1,1,2) & "' AND FECHA_SALIDAS <= '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S1T, " &_
								"ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
								"INTO #SALIDAS  " &_
								"FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
								"LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO= A.CONT_SAE " &_  
								"LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO= C.SIGLA  " &_
								"WHERE A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
								"AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES', 'PYME')  " &_
								"AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
								"AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z','W') " &_ 
								"GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
								"ltrim(rtrim(A.SEGMENTO_RIESGO)) " 

								''response.write sql
					        	conn.execute sql	
					    else
					    	if f2<>"" then
					        sql= "SELECT " &_
								"IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio, " &_
								"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') and  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAA, " &_ 
								"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAAConVenta,   " &_ 
								"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as Salidas, " &_
								"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' THEN soles + dolares*tipocambio ELSE 0 END) as SalidaConVenta, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  FECHA_SALIDAS <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S4T, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' AND FECHA_SALIDAS <= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S3T, " &_
								"SUM(CASE WHEN FECHA_SALIDAS > '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' AND FECHA_SALIDAS <= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S2T, " &_
								"ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
								"INTO #SALIDAS  " &_
								"FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
								"LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO= A.CONT_SAE " &_  
								"LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO= C.SIGLA  " &_
								"WHERE A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
								"AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES', 'PYME')  " &_
								"AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
								"AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z','W') " &_ 
								"GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
								"ltrim(rtrim(A.SEGMENTO_RIESGO)) " 

								''response.write sql
					        	conn.execute sql	
						    else
						    	if f3<>"" then
							        sql= "SELECT " &_
										"IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio, " &_
										"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') and  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAA, " &_ 
										"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAAConVenta,   " &_ 
										"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as Salidas, " &_
										"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierremesant,7,4) & mid(fechacierremesant,4,2) & mid(fechacierremesant,1,2) & "' THEN soles + dolares*tipocambio ELSE 0 END) as SalidaConVenta, " &_
										"SUM(CASE WHEN FECHA_SALIDAS >'" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  FECHA_SALIDAS <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S4T, " &_
								        "SUM(CASE WHEN FECHA_SALIDAS >'" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' AND FECHA_SALIDAS <= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S3T, " &_
								        "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
										"INTO #SALIDAS  " &_
										"FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
										"LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO= A.CONT_SAE " &_  
										"LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO= C.SIGLA  " &_
										"WHERE A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
										"AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES', 'PYME')  " &_
										"AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
										"AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z','W') " &_ 
										"GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
										"ltrim(rtrim(A.SEGMENTO_RIESGO)) " 

										''response.write sql
							        	conn.execute sql	
							    else
							    	if f4<>"" then
							        sql= "SELECT " &_
										"IsNull(C.nomb_estu,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS as estudio, " &_
										"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') and  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAA, " &_ 
										"SUM(CASE WHEN YEAR(A.FECHA_SALIDAS)=year('" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') THEN soles + dolares*tipocambio ELSE 0 END) as SalidasAAAAConVenta,   " &_ 
										"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as Salidas, " &_
										"SUM(CASE WHEN FECHA_SALIDAS>'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' THEN soles + dolares*tipocambio ELSE 0 END) as SalidaConVenta, " &_
										"SUM(CASE WHEN FECHA_SALIDAS> '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' AND  FECHA_SALIDAS <= '" & mid(f5,7,4) & mid(f5,4,2) & mid(f5,1,2) & "' AND  MODALIDAD NOT IN ('W') THEN soles + dolares*tipocambio ELSE 0 END) as S4T, " &_
								        "ltrim(rtrim(A.SEGMENTO_RIESGO)) as SEGMENTO_RIESGO " &_
										"INTO #SALIDAS  " &_
										"FROM FB.CentroCobranzas.dbo.Vista_Salidas_Segmento A  " &_
										"LEFT JOIN	FB.CentroCobranzas.dbo.PD_ULT_CONTRATO_ESTUDIO M	ON M.CONTRATO= A.CONT_SAE " &_  
										"LEFT JOIN	FB.CentroCobranzas.dbo.PD_CATALOGO_ESTUDIOS C	ON M.CODDATO= C.SIGLA  " &_
										"WHERE A.FECHA_SALIDAS<='" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "' " &_
										"AND RTRIM(LTRIM(A.SEGMENTO_RIESGO))	in ('PARTICULARES', 'PYME')  " &_
										"AND ltrim(rtrim(A.entidad))			in ('BC','35')	 " &_
										"AND	A.MODALIDAD	in ('B','D','E','G','M','N','R','S','T','Y','Z','W') " &_ 
										"GROUP BY IsNull(C.nomb_estu,'NO IDENTIFICADO'), " &_
										"ltrim(rtrim(A.SEGMENTO_RIESGO)) " 

										''response.write sql
							        	conn.execute sql	
								    end if
							    end if
						    end if
					    end if
				        
					    if f1<>"" then
					    	sql2="Select IsNull(A.estudio,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS AS Estudio, " &_
					    	"IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f1,7,4) & mid(f1,4,2) & mid(f1,1,2) & "' then A.JUD else 0 end),0) as StockCierrePART1, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f1,7,4) & mid(f1,4,2) & mid(f1,1,2) & "' then A.JUD else 0 end),0) as StockCierrePYME1, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' then JUD else 0 end),0) as StockCierrePART2, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' then JUD else 0 end),0) as StockCierrePYME2 " &_
			                            "IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePART4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePYME4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' then JUD else 0 end),0) as StockCierrePART3, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' then JUD else 0 end),0) as StockCierrePYME3 " &_
			                        	"INTO #STOCKS " &_
			                        	"from FB.CentroCobranzas.dbo.PD_Casos_Jud A where SEGMENTO_RIESGO in ('PARTICULARES','PYME') and fasigna is not null group by estudio "

			                      conn.execute sql2


						    	sql="select estudio, sum (case when segmento_riesgo='particulares' then 1 else 0 end) AS NUMCASOSPART, " &_
	                        	"sum (case when segmento_riesgo='PYME' then 1 else 0 end) AS NUMCASOSPYME, " &_
	                        	"count (distinct (case when segmento_riesgo='PARTICULARES' then codcentral end)) AS NUMCLIENTESPART, " &_
	                        	"count (distinct (case when segmento_riesgo='PYME' then codcentral end)) AS NUMCLIENTESPYME, " &_
	                        	"sum(case when segmento_riesgo='PARTICULARES' then PROVCONST else 0 end) as PROVCONSTPART, " &_
	                        	"sum(case when segmento_riesgo='PYME' then PROVCONST else 0 end) as PROVCONSTPYME, " &_
	                        	"SUM(case when segmento_riesgo='PARTICULARES' then JUD else 0 end) as StockCierrePART, " &_
	                        	"SUM(case when segmento_riesgo='PYME' then JUD else 0 end) as StockCierrePYME, " &_
	                        	"IsNull((Select StockCierrePART4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART4," &_
	                        	"IsNull((Select StockCierrePART3 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART3," &_
	                        	"IsNull((Select StockCierrePART2 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART2," &_
	                        	"IsNull((Select StockCierrePART1 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART1," &_
	                        	"IsNull((Select StockCierrePYME4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME4," &_
	                        	"IsNull((Select StockCierrePYME3 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME3," &_
	                        	"IsNull((Select StockCierrePYME2 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME2," &_
	                        	"IsNull((Select StockCierrePYME1 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME1," &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as INGRESOSPART, " &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as INGRESOSPYME, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASPART, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASPYME, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASCONVENTAPART, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASCONVENTAPYME, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAPART, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAPYME, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAConVentaPART, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAConVentaPYME, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT4PART, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT4PYME, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S4TPART, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S4TPYME, " &_
	                            "IsNull((select sum(IT3) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT3PART, " &_
	                            "IsNull((select sum(IT3) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT3PYME, " &_
	                            "IsNull((select sum(S3T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S3TPART, " &_
	                            "IsNull((select sum(S3T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S3TPYME " &_
	                            "IsNull((select sum(IT2) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT2PART, " &_
	                            "IsNull((select sum(IT2) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT2PYME, " &_
	                            "IsNull((select sum(S2T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S2TPART, " &_
	                            "IsNull((select sum(S2T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S2TPYME, " &_
	                            "IsNull((select sum(IT1) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT1PART, " &_
	                            "IsNull((select sum(IT1) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT1PYME, " &_
	                            "IsNull((select sum(S1T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S1TPART, " &_
	                            "IsNull((select sum(S1T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S1TPYME " &_
	                            "from FB.CentroCobranzas.dbo.PD_Casos_Jud A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and A.fasigna is not null AND ltrim(rtrim(A.Plaza)) = '" & plaza & "' and fechadatos=(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') group by estudio order by (StockCierrePART+StockCierrePYME) desc"

					    else

					    	if f2<>"" then
						    	sql2="Select IsNull(A.estudio,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS AS Estudio, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' then JUD else 0 end),0) as StockCierrePART2, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f2,7,4) & mid(f2,4,2) & mid(f2,1,2) & "' then JUD else 0 end),0) as StockCierrePYME2 " &_
			                            "IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePART4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePYME4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' then JUD else 0 end),0) as StockCierrePART3, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' then JUD else 0 end),0) as StockCierrePYME3 " &_
			                        	"INTO #STOCKS " &_
			                        	"from FB.CentroCobranzas.dbo.PD_Casos_Jud A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and fasigna is not null group by estudio "

			                      conn.execute sql2


						    	sql="select estudio, sum (case when segmento_riesgo='particulares' then 1 else 0 end) AS NUMCASOSPART, " &_
	                        	"sum (case when segmento_riesgo='PYME' then 1 else 0 end) AS NUMCASOSPYME, " &_
	                        	"count (distinct (case when segmento_riesgo='PARTICULARES' then codcentral end)) AS NUMCLIENTESPART, " &_
	                        	"count (distinct (case when segmento_riesgo='PYME' then codcentral end)) AS NUMCLIENTESPYME, " &_
	                        	"sum(case when segmento_riesgo='PARTICULARES' then PROVCONST else 0 end) as PROVCONSTPART, " &_
	                        	"sum(case when segmento_riesgo='PYME' then PROVCONST else 0 end) as PROVCONSTPYME, " &_
	                        	"SUM(case when segmento_riesgo='PARTICULARES' then JUD else 0 end) as StockCierrePART, " &_
	                        	"SUM(case when segmento_riesgo='PYME' then JUD else 0 end) as StockCierrePYME, " &_
	                        	"IsNull((Select StockCierrePART4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART4," &_
	                        	"IsNull((Select StockCierrePART3 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART3," &_
	                        	"IsNull((Select StockCierrePART2 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART2," &_
	                        	"IsNull((Select StockCierrePYME4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME4," &_
	                        	"IsNull((Select StockCierrePYME3 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME3," &_
	                        	"IsNull((Select StockCierrePYME2 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME2," &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as INGRESOSPART, " &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as INGRESOSPYME, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASPART, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASPYME, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASCONVENTAPART, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASCONVENTAPYME, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAPART, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAPYME, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAConVentaPART, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAConVentaPYME, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT4PART, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT4PYME, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S4TPART, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S4TPYME, " &_
	                            "IsNull((select sum(IT3) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT3PART, " &_
	                            "IsNull((select sum(IT3) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT3PYME, " &_
	                            "IsNull((select sum(S3T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S3TPART, " &_
	                            "IsNull((select sum(S3T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S3TPYME " &_
	                            "IsNull((select sum(IT2) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT2PART, " &_
	                            "IsNull((select sum(IT2) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT2PYME, " &_
	                            "IsNull((select sum(S2T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S2TPART, " &_
	                            "IsNull((select sum(S2T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S2TPYME, " &_
	                            "from FB.CentroCobranzas.dbo.PD_Casos_Jud A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and A.fasigna is not null AND ltrim(rtrim(A.Plaza)) = '" & plaza & "' and fechadatos=(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') group by estudio order by (StockCierrePART+StockCierrePYME) desc"

					    else

					    	if f3<>"" then

					    	sql2="Select IsNull(A.estudio,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS AS Estudio, " &_
			                            "IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePART4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePYME4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' then JUD else 0 end),0) as StockCierrePART3, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f3,7,4) & mid(f3,4,2) & mid(f3,1,2) & "' then JUD else 0 end),0) as StockCierrePYME3 " &_
			                        	"INTO #STOCKS " &_
			                        	"from FB.CentroCobranzas.dbo.PD_Casos_Jud A where SEGMENTO_RIESGO in ('PARTICULARES','PYME') and fasigna is not null group by estudio "

			                      conn.execute sql2


						    	sql="select estudio, sum (case when segmento_riesgo='particulares' then 1 else 0 end) AS NUMCASOSPART, " &_
	                        	"sum (case when segmento_riesgo='PYME' then 1 else 0 end) AS NUMCASOSPYME, " &_
	                        	"count (distinct (case when segmento_riesgo='PARTICULARES' then codcentral end)) AS NUMCLIENTESPART, " &_
	                        	"count (distinct (case when segmento_riesgo='PYME' then codcentral end)) AS NUMCLIENTESPYME, " &_
	                        	"sum(case when segmento_riesgo='PARTICULARES' then PROVCONST else 0 end) as PROVCONSTPART, " &_
	                        	"sum(case when segmento_riesgo='PYME' then PROVCONST else 0 end) as PROVCONSTPYME, " &_
	                        	"SUM(case when segmento_riesgo='PARTICULARES' then JUD else 0 end) as StockCierrePART, " &_
	                        	"SUM(case when segmento_riesgo='PYME' then JUD else 0 end) as StockCierrePYME, " &_
	                        	"IsNull((Select StockCierrePART4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART4," &_
	                        	"IsNull((Select StockCierrePART3 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART3," &_
	                        	"IsNull((Select StockCierrePYME4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME4," &_
	                        	"IsNull((Select StockCierrePYME3 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME3," &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as INGRESOSPART, " &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as INGRESOSPYME, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASPART, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASPYME, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASCONVENTAPART, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASCONVENTAPYME, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAPART, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAPYME, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAConVentaPART, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAConVentaPYME, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT4PART, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT4PYME, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S4TPART, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S4TPYME, " &_
	                            "IsNull((select sum(IT3) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT3PART, " &_
	                            "IsNull((select sum(IT3) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT3PYME, " &_
	                            "IsNull((select sum(S3T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S3TPART, " &_
	                            "IsNull((select sum(S3T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S3TPYME " &_
	                            "from FB.CentroCobranzas.dbo.PD_Casos_Jud A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and A.fasigna is not null AND ltrim(rtrim(A.Plaza)) = '" & plaza & "' and fechadatos=(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') group by estudio order by StockCierrePART desc"

						    else
						    	if f4<>"" then
			                        sql2="Select IsNull(A.estudio,'NO IDENTIFICADO') Collate SQL_Latin1_General_CP1_CI_AS AS Estudio, " &_
			                            "IsNull(SUM(case when A.segmento_riesgo='PARTICULARES' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePART4, " &_
			                        	"IsNull(SUM(case when A.segmento_riesgo='PYME' AND A.fechadatos= '" & mid(f4,7,4) & mid(f4,4,2) & mid(f4,1,2) & "' then A.JUD else 0 end),0) as StockCierrePYME4, " &_
			                        	"INTO #STOCKS " &_
			                        	"from FB.CentroCobranzas.dbo.PD_Casos_Jud A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and fasigna is not null group by estudio "

			                      conn.execute sql2


						    	sql="select estudio, sum (case when segmento_riesgo='particulares' then 1 else 0 end) AS NUMCASOSPART, " &_
	                        	"sum (case when segmento_riesgo='PYME' then 1 else 0 end) AS NUMCASOSPYME, " &_
	                        	"count (distinct (case when segmento_riesgo='PARTICULARES' then codcentral end)) AS NUMCLIENTESPART, " &_
	                        	"count (distinct (case when segmento_riesgo='PYME' then codcentral end)) AS NUMCLIENTESPYME, " &_
	                        	"sum(case when segmento_riesgo='PARTICULARES' then PROVCONST else 0 end) as PROVCONSTPART, " &_
	                        	"sum(case when segmento_riesgo='PYME' then PROVCONST else 0 end) as PROVCONSTPYME, " &_
	                        	"SUM(case when segmento_riesgo='PARTICULARES' then JUD else 0 end) as StockCierrePART, " &_
	                        	"SUM(case when segmento_riesgo='PYME' then JUD else 0 end) as StockCierrePYME, " &_
	                        	"IsNull((Select StockCierrePART4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePART4," &_
	                        	"IsNull((Select StockCierrePYME4 from #STOCKS where estudio=A.estudio),0) AS StockCierrePYME4," &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as INGRESOSPART, " &_
	                            "IsNull((select sum(ingresos) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as INGRESOSPYME, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASPART, " &_
	                            "IsNull((select sum(Salidas) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASPYME, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SALIDASCONVENTAPART, " &_
	                            "IsNull((select sum(salidaconventa) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SALIDASCONVENTAPYME, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAPART, " &_
	                            "IsNull((select sum(SalidasAAAA) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAPYME, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as SalidasAAAAConVentaPART, " &_
	                            "IsNull((select sum(SalidasAAAAConVenta) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as SalidasAAAAConVentaPYME, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as IT4PART, " &_
	                            "IsNull((select sum(IT4) from #INGRESOS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as IT4PYME, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PARTICULARES'),0) as S4TPART, " &_
	                            "IsNull((select sum(S4T) from #SALIDAS where estudio=A.estudio and SEGMENTO_RIESGO='PYME'),0) as S4TPYME, " &_
	                            "from FB.CentroCobranzas.dbo.PD_Casos_Jud A where A.SEGMENTO_RIESGO in ('PARTICULARES','PYME') and A.fasigna is not null AND ltrim(rtrim(A.Plaza)) = '" & plaza & "' and fechadatos=(select MAX(fechacierre) from FB.CentroCobranzas.dbo.PD_FechasCierre where fechacierre<'" & mid(fechacierre,7,4) & mid(fechacierre,4,2) & mid(fechacierre,1,2) & "') group by estudio order by (StockCierrePART+StockCierrePYME) desc"

			                        ''response.Write sql
			                    end if
						    end if
					    end if
					end if 

					    
	                    consultar sql,RS	
						bgcolor="#FFFFFF"

						SUMA_NUMCASOSPART=0
                        SUMA_NUMCASOSPYME=0
                        SUMA_NUMCLIENTESPART=0
                        SUMA_NUMCLIENTESPYME=0
                        SUMA_PROVCONSTPART=0
                        SUMA_PROVCONSTPYME=0
                        SUMA_ASIGNADOPART=0
                        SUMA_ASIGNADOPYME=0
                        SUMA_SALIDASSINVENTAAAAPART=0
                        SUMA_SALIDASCONVENTAAAAPART=0
                        SUMA_SALIDASSINVENTMESPART=0
                        SUMA_SALIDASCONVENTMESPART=0
                        SUMA_SALIDASSINVENTAAAAPYME=0
                        SUMA_SALIDASCONVENTAAAAPYME=0
                        SUMA_SALIDASSINVENTMESPYME=0
                        SUMA_SALIDASCONVENTMESPYME=0
                        SUMA_META=0
                        SUMA_S4TPART=0
                        SUMA_S4TPYME=0
                        SUMA_S3TPART=0
                        SUMA_S3TPYME=0
                        SUMA_S2TPART=0
                        SUMA_S2TPYME=0
                        SUMA_S1TPART=0
                        SUMA_S1TPYME=0
                        SUMA_IT4PYME=0
                        SUMA_IT4PART=0
                        SUMA_IT3PYME=0
                        SUMA_IT3PART=0
                        SUMA_IT2PYME=0
                        SUMA_IT2PART=0
                        SUMA_IT1PYME=0
                        SUMA_IT1PART=0
                        SUMA_StockCierrePYME4=0
                        SUMA_StockCierrePART4=0
                        SUMA_StockCierrePYME3=0
                        SUMA_StockCierrePART3=0
                        SUMA_StockCierrePYME2=0
                        SUMA_StockCierrePART2=0
                        SUMA_StockCierrePYME1=0
                        SUMA_StockCierrePART1=0

                        Do While not RS.EOF

                        	SUMA_NUMCASOSPART= SUMA_NUMCASOSPART + RS.Fields("NUMCASOSPART")
	                        SUMA_NUMCASOSPYME=SUMA_NUMCASOSPYME + RS.Fields("NUMCASOSPYME")
	                        SUMA_NUMCLIENTESPART=SUMA_NUMCLIENTESPART + RS.Fields("NUMCLIENTESPART")
	                        SUMA_NUMCLIENTESPYME=SUMA_NUMCLIENTESPYME + RS.Fields("NUMCLIENTESPYME")
	                        SUMA_PROVCONSTPART=SUMA_PROVCONSTPART + RS.Fields("PROVCONSTPART")/1000
	                        SUMA_PROVCONSTPYME=SUMA_PROVCONSTPYME + RS.Fields("PROVCONSTPYME")/1000
	                        SUMA_ASIGNADOPART=SUMA_ASIGNADOPART + RS.Fields("StockCierrePART")/1000
	                        SUMA_ASIGNADOPYME=SUMA_ASIGNADOPYME + RS.Fields("StockCierrePYME")/1000
	                        SUMA_SALIDASSINVENTAAAAPART=SUMA_SALIDASSINVENTAAAAPART + RS.Fields("SalidasAAAAPART")/1000
	                        SUMA_SALIDASCONVENTAAAAPART=SUMA_SALIDASCONVENTAAAAPART + RS.Fields("SalidasAAAAConVentaPART")/1000
	                        SUMA_SALIDASSINVENTMESPART=SUMA_SALIDASSINVENTMESPART +  RS.Fields("SALIDASPART")/1000
	                        SUMA_SALIDASCONVENTMESPART=SUMA_SALIDASCONVENTMESPART +  RS.Fields("SALIDASCONVENTAPART")/1000
	                        SUMA_SALIDASSINVENTAAAAPYME=SUMA_SALIDASSINVENTAAAAPYME + RS.Fields("SalidasAAAAPYME")/1000
	                        SUMA_SALIDASCONVENTAAAAPYME=SUMA_SALIDASCONVENTAAAAPYME + RS.Fields("SalidasAAAAConVentaPYME")/1000
	                        SUMA_SALIDASSINVENTMESPYME=SUMA_SALIDASSINVENTMESPYME +  RS.Fields("SALIDASPYME")/1000
	                        SUMA_SALIDASCONVENTMESPYME=SUMA_SALIDASCONVENTMESPYME +  RS.Fields("SALIDASCONVENTAPYME")/1000
	                        SUMA_META=SUMA_META +  0
	                        if f1 <>"" then
	                            SUMA_S4TPART=SUMA_S4TPART +  RS.Fields("S4TPART")/1000
                        	    SUMA_S4TPYME=SUMA_S4TPYME +  RS.Fields("S4TPYME")/1000
                        	    SUMA_S3TPART=SUMA_S3TPART +  RS.Fields("S3TPART")/1000
                        	    SUMA_S3TPYME=SUMA_S3TPYME +  RS.Fields("S3TPYME")/1000
                        	    SUMA_S2TPART=SUMA_S2TPART +  RS.Fields("S2TPART")/1000
                        	    SUMA_S2TPYME=SUMA_S2TPYME +  RS.Fields("S2TPYME")/1000
                        	    SUMA_S1TPART=SUMA_S1TPART +  RS.Fields("S1TPART")/1000
                        	    SUMA_S1TPYME=SUMA_S1TPYME +  RS.Fields("S1TPYME")/1000
                        	    SUMA_IT4PART=SUMA_IT4PART +  RS.Fields("IT4PART")/1000
                        	    SUMA_IT3PART=SUMA_IT3PART +  RS.Fields("IT3PART")/1000
                        	    SUMA_IT4PYME=SUMA_IT4PYME +  RS.Fields("IT4PYME")/1000
                        	    SUMA_IT3PYME=SUMA_IT3PYME +  RS.Fields("IT3PYME")/1000
                        	    SUMA_IT2PART=SUMA_IT2PART +  RS.Fields("IT2PART")/1000
                        	    SUMA_IT1PART=SUMA_IT1PART +  RS.Fields("IT1PART")/1000
                        	    SUMA_IT2PYME=SUMA_IT2PYME +  RS.Fields("IT2PYME")/1000
                        	    SUMA_IT1PYME=SUMA_IT1PYME +  RS.Fields("IT1PYME")/1000
                        	    SUMA_StockCierrePYME4=SUMA_StockCierrePYME4  +  RS.Fields("StockCierrePYME4")/1000
                        	    SUMA_StockCierrePART4=SUMA_StockCierrePART4  +  RS.Fields("StockCierrePART4")/1000
                        	    SUMA_StockCierrePYME3=SUMA_StockCierrePYME3  +  RS.Fields("StockCierrePYME3")/1000
                        	    SUMA_StockCierrePART3=SUMA_StockCierrePART3  +  RS.Fields("StockCierrePART3")/1000
                        	    SUMA_StockCierrePYME2=SUMA_StockCierrePYME2  +  RS.Fields("StockCierrePYME2")/1000
                        	    SUMA_StockCierrePART2=SUMA_StockCierrePART2  +  RS.Fields("StockCierrePART2")/1000
                        	    SUMA_StockCierrePYME1=SUMA_StockCierrePYME1  +  RS.Fields("StockCierrePYME1")/1000
                        	    SUMA_StockCierrePART1=SUMA_StockCierrePART1  +  RS.Fields("StockCierrePART1")/1000
                        	 else
                        	    if f2 <>"" then
	                                SUMA_S4TPART=SUMA_S4TPART +  RS.Fields("S4TPART")/1000
                        	        SUMA_S4TPYME=SUMA_S4TPYME +  RS.Fields("S4TPYME")/1000
                        	        SUMA_S3TPART=SUMA_S3TPART +  RS.Fields("S3TPART")/1000
                        	        SUMA_S3TPYME=SUMA_S3TPYME +  RS.Fields("S3TPYME")/1000
                        	        SUMA_S2TPART=SUMA_S2TPART +  RS.Fields("S2TPART")/1000
                        	        SUMA_S2TPYME=SUMA_S2TPYME +  RS.Fields("S2TPYME")/1000
                        	        SUMA_IT4PART=SUMA_IT4PART +  RS.Fields("IT4PART")/1000
                        	        SUMA_IT3PART=SUMA_IT3PART +  RS.Fields("IT3PART")/1000
                        	        SUMA_IT4PYME=SUMA_IT4PYME +  RS.Fields("IT4PYME")/1000
                        	        SUMA_IT3PYME=SUMA_IT3PYME +  RS.Fields("IT3PYME")/1000
                        	        SUMA_IT2PART=SUMA_IT2PART +  RS.Fields("IT2PART")/1000
                        	        SUMA_IT2PYME=SUMA_IT2PYME +  RS.Fields("IT2PYME")/1000
                        	        SUMA_StockCierrePYME4=SUMA_StockCierrePYME4  +  RS.Fields("StockCierrePYME4")/1000
                        	        SUMA_StockCierrePART4=SUMA_StockCierrePART4  +  RS.Fields("StockCierrePART4")/1000
                        	        SUMA_StockCierrePYME3=SUMA_StockCierrePYME3  +  RS.Fields("StockCierrePYME3")/1000
                        	        SUMA_StockCierrePART3=SUMA_StockCierrePART3  +  RS.Fields("StockCierrePART3")/1000
                        	        SUMA_StockCierrePYME2=SUMA_StockCierrePYME2  +  RS.Fields("StockCierrePYME2")/1000
                        	        SUMA_StockCierrePART2=SUMA_StockCierrePART2  +  RS.Fields("StockCierrePART2")/1000
                        	     else
                        	        if f3 <>"" then
	                                    SUMA_S4TPART=SUMA_S4TPART +  RS.Fields("S4TPART")/1000
                        	            SUMA_S4TPYME=SUMA_S4TPYME +  RS.Fields("S4TPYME")/1000
                        	            SUMA_S3TPART=SUMA_S3TPART +  RS.Fields("S3TPART")/1000
                        	            SUMA_S3TPYME=SUMA_S3TPYME +  RS.Fields("S3TPYME")/1000
                        	            SUMA_IT4PART=SUMA_IT4PART +  RS.Fields("IT4PART")/1000
                        	            SUMA_IT3PART=SUMA_IT3PART +  RS.Fields("IT3PART")/1000
                        	            SUMA_IT4PYME=SUMA_IT4PYME +  RS.Fields("IT4PYME")/1000
                        	            SUMA_IT3PYME=SUMA_IT3PYME +  RS.Fields("IT3PYME")/1000
                        	            SUMA_StockCierrePYME4=SUMA_StockCierrePYME4  +  RS.Fields("StockCierrePYME4")/1000
                        	            SUMA_StockCierrePART4=SUMA_StockCierrePART4  +  RS.Fields("StockCierrePART4")/1000
                        	            SUMA_StockCierrePYME3=SUMA_StockCierrePYME3  +  RS.Fields("StockCierrePYME3")/1000
                        	            SUMA_StockCierrePART3=SUMA_StockCierrePART3  +  RS.Fields("StockCierrePART3")/1000
                        	         else
                                	    if f4 <>"" then
	                                        SUMA_S4TPART=SUMA_S4TPART +  RS.Fields("S4TPART")/1000
                        	                SUMA_S4TPYME=SUMA_S4TPYME +  RS.Fields("S4TPYME")/1000
                        	                SUMA_IT4PART=SUMA_IT4PART +  RS.Fields("IT4PART")/1000
                        	                SUMA_IT4PYME=SUMA_IT4PYME +  RS.Fields("IT4PYME")/1000
                        	                SUMA_StockCierrePYME4=SUMA_StockCierrePYME4  +  RS.Fields("StockCierrePYME4")/1000
                        	                SUMA_StockCierrePART4=SUMA_StockCierrePART4  +  RS.Fields("StockCierrePART4")/1000
                        	             else
                                    	    
                                    	 
                        	             end if
                                	 
                        	         end if
                            	    
                            	 
                        	     end if
                        	 
                        	 end if
                            
                        RS.MoveNext
                        Loop
                        
                        RS.MoveFirst
                        %>

                        <tr>
						    <td bgcolor="#BEE8FB">
							    <font size=1 color=#483d8b face=Arial><b>TOTAL</b></font>
						    </td>
						    <td bgcolor="#BEE8FB">
							    <font size=1 color=#483d8b face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_NUMCLIENTESPART+SUMA_NUMCLIENTESPYME,0)%></b></font>
						    </td>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_NUMCASOSPART+SUMA_NUMCASOSPYME,0)%></b></font>
						    </td>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_PROVCONSTPART+SUMA_PROVCONSTPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_ASIGNADOPART+SUMA_ASIGNADOPYME,0)%></b></font>
						    </td>	
						    <td bgcolor="#FFFFFF"  align="center" width="1%">
								<font size=1 color=#483d8b face=Arial><b>&nbsp;</b></font>
							</td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASSINVENTAAAAPART+SUMA_SALIDASSINVENTAAAAPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASCONVENTAAAAPYME+SUMA_SALIDASCONVENTAAAAPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASSINVENTMESPYME+SUMA_SALIDASSINVENTMESPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASCONVENTMESPYME+SUMA_SALIDASCONVENTMESPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META,0)%></b></font>
						    </td>
						    <%if f1<>"" then %>	
						   									
						        <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S1TPART+SUMA_S1TPYME)>0,SUMA_S1TPART+SUMA_S1TPYME,1),0)%>%</b></font>
						        </td>
						   	    <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							    <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S2TPART+SUMA_S2TPYME)>0,SUMA_S2TPART+SUMA_S2TPYME,1),0)%>%</b></font>
						        </td>
						   	    <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							    <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPART+SUMA_S3TPYME)>0,SUMA_S3TPART+SUMA_S3TPYME,1),0)%>%</b></font>
						        </td>
						   	    <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							    <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART+SUMA_S4TPYME)>0,SUMA_S4TPART+SUMA_S4TPYME,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
    						    
						        
							 <%else %>
							        <%if f2<>"" then %>	
    						   									
    						            
							            <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S2TPART+SUMA_S2TPYME)>0,SUMA_S2TPART+SUMA_S2TPYME,1),0)%>%</b></font>
						                </td>
						   	            <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
							            <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPART+SUMA_S3TPYME)>0,SUMA_S3TPART+SUMA_S3TPYME,1),0)%>%</b></font>
						                </td>
						   	            <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
							            <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART+SUMA_S4TPYME)>0,SUMA_S4TPART+SUMA_S4TPYME,1),0)%>%</b></font>
						                </td>
						                <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
							     <%else %>
							        <%if f3<>"" then %>	

							            <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPART+SUMA_S3TPYME)>0,SUMA_S3TPART+SUMA_S3TPYME,1),0)%>%</b></font>
						                </td>
						   	            <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
							            <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART+SUMA_S4TPYME)>0,SUMA_S4TPART+SUMA_S4TPYME,1),0)%>%</b></font>
						                </td>
						                <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
							     <%else %>
							        <%if f4<>"" then %>	

							                <td bgcolor="#BEE8FB" align="right" width="5%">
							                    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART+SUMA_S4TPYME)>0,SUMA_S4TPART+SUMA_S4TPYME,1),0)%>%</b></font>
						                    </td>
						                    <td bgcolor="#FFFFFF" align="center" width="2%">
								                <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							                </td>

							        
							            <%end if %>
							        <%end if %>
							     <%end if %>
							 <%end if %>							
						   									
						    
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*12/iif((SUMA_SALIDASSINVENTAAAAPYME+SUMA_SALIDASSINVENTAAAAPART)>0,SUMA_SALIDASSINVENTAAAAPYME+SUMA_SALIDASSINVENTAAAAPART,1),0)%>%</b></font>
						    </td>
						  
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber((SUMA_S4TPART+SUMA_S4TPYME)*100/iif((SUMA_StockCierrePART4+SUMA_StockCierrePYME4+SUMA_IT4PYME+SUMA_IT4PART)>0,SUMA_StockCierrePART4+SUMA_StockCierrePYME4+SUMA_IT4PYME+SUMA_IT4PART,1),2)%>%</b></font>
						    </td>
						    
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber((SUMA_S3TPART+SUMA_S3TPYME)*100/iif((SUMA_StockCierrePART3+SUMA_StockCierrePYME3+SUMA_IT3PYME+SUMA_IT3PART)>0,SUMA_StockCierrePART3+SUMA_StockCierrePYME3+SUMA_IT3PYME+SUMA_IT3PART,1),2)%>%</b></font>
						    </td>
	                    </tr>

	                    <tr>
						    <td bgcolor="#BEE8FB" rowspan="2">
							    <font size=1 color=#483d8b face=Arial><b><%=Plaza%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB">
							    <font size=1 color=#483d8b face=Arial><b>PARTICULARES</b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_NUMCLIENTESPART,0)%></b></font>
						    </td>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_NUMCASOSPART,0)%></b></font>
						    </td>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_PROVCONSTPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_ASIGNADOPART,0)%></b></font>
						    </td>	
						    <td bgcolor="#FFFFFF"  align="center" width="1%">
								<font size=1 color=#483d8b face=Arial><b>&nbsp;</b></font>
							</td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASSINVENTAAAAPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASCONVENTAAAAPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASSINVENTMESPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASCONVENTMESPART,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META,0)%></b></font>
						    </td>	
				            <%if f1<>"" then%>
				                <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART)>0,SUMA_S4TPART,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
    						    
						        <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPART)>0,SUMA_S3TPART,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF"  align="center" width="2%">
								    <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							    <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S2TPART)>0,SUMA_S2TPART,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF"  align="center" width="2%">
								    <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							    <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S1TPART)>0,SUMA_S1TPART,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF"  align="center" width="2%">
								    <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
				            <%else %>
				                <%if f2<>"" then%>
				                    <td bgcolor="#BEE8FB" align="right" width="5%">
							            <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART)>0,SUMA_S4TPART,1),0)%>%</b></font>
						            </td>
						            <td bgcolor="#FFFFFF" align="center" width="2%">
								        <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							        </td>
        						    
						            <td bgcolor="#BEE8FB" align="right" width="5%">
							            <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPART)>0,SUMA_S3TPART,1),0)%>%</b></font>
						            </td>
						            <td bgcolor="#FFFFFF"  align="center" width="2%">
								        <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							        </td>
							        <td bgcolor="#BEE8FB" align="right" width="5%">
							            <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S2TPART)>0,SUMA_S2TPART,1),0)%>%</b></font>
						            </td>
						            <td bgcolor="#FFFFFF"  align="center" width="2%">
								        <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							        </td>
				                <%else %>
				                    <%if f3<>"" then%>
				                        <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART)>0,SUMA_S4TPART,1),0)%>%</b></font>
						                </td>
						                <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
            						    
						                <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPART)>0,SUMA_S3TPART,1),0)%>%</b></font>
						                </td>
						                <td bgcolor="#FFFFFF"  align="center" width="2%">
								            <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>

				                    <%else %>
				                        <%if f4<>"" then%>
				                            <td bgcolor="#BEE8FB" align="right" width="5%">
							                    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPART)>0,SUMA_S4TPART,1),0)%>%</b></font>
						                    </td>
						                    <td bgcolor="#FFFFFF" align="center" width="2%">
								                <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							                </td>


				                        <%else %>
				                        <%end if %>
				                    <%end if %>
				                <%end if %>
				            <%end if %>
						    
							
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*12/iif((SUMA_SALIDASSINVENTAAAAPART)>0,SUMA_SALIDASSINVENTAAAAPART,1),0)%>%</b></font>
						    </td>
						  
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber((SUMA_S4TPART)*100/iif((SUMA_StockCierrePART4+SUMA_IT4PART)>0,SUMA_StockCierrePART4+SUMA_IT4PART,1),2)%>%</b></font>
						    </td>
						    
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber((SUMA_S3TPART)*100/iif((SUMA_StockCierrePART3+SUMA_IT3PART)>0,SUMA_StockCierrePART3+SUMA_IT3PART,1),2)%>%</b></font>
						    </td>
	                    </tr>
	                    <tr>
						    
						    <td bgcolor="#BEE8FB">
							    <font size=1 color=#483d8b face=Arial><b>PYME</b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_NUMCLIENTESPYME,0)%></b></font>
						    </td>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_NUMCASOSPYME,0)%></b></font>
						    </td>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_PROVCONSTPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_ASIGNADOPYME,0)%></b></font>
						    </td>	
						    <td bgcolor="#FFFFFF"  align="center" width="1%">
								<font size=1 color=#483d8b face=Arial><b>&nbsp;</b></font>
							</td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASSINVENTAAAAPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASCONVENTAAAAPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASSINVENTMESPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_SALIDASCONVENTMESPYME,0)%></b></font>
						    </td>
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META,0)%></b></font>
						    </td>	
									
						    <%if f1<>"" then %>
						        <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPYME)>0,SUMA_S4TPYME,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
    						    
						        <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPYME)>0,SUMA_S3TPYME,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF"  align="center" width="2%">
								    <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							    <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S2TPYME)>0,SUMA_S2TPYME,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
    						    
						        <td bgcolor="#BEE8FB" align="right" width="5%">
							        <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S1TPYME)>0,SUMA_S1TPYME,1),0)%>%</b></font>
						        </td>
						        <td bgcolor="#FFFFFF"  align="center" width="2%">
								    <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							    </td>
							<%else %>
							    <%if f2<>"" then%>
						            <td bgcolor="#BEE8FB" align="right" width="5%">
							            <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPYME)>0,SUMA_S4TPYME,1),0)%>%</b></font>
						            </td>
						            <td bgcolor="#FFFFFF" align="center" width="2%">
								        <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							        </td>
        						    
						            <td bgcolor="#BEE8FB" align="right" width="5%">
							            <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPYME)>0,SUMA_S3TPYME,1),0)%>%</b></font>
						            </td>
						            <td bgcolor="#FFFFFF"  align="center" width="2%">
								        <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							        </td>
							        <td bgcolor="#BEE8FB" align="right" width="5%">
							            <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S2TPYME)>0,SUMA_S2TPYME,1),0)%>%</b></font>
						            </td>
						            <td bgcolor="#FFFFFF" align="center" width="2%">
								        <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							        </td>

							    <%else %>
							        <%if f3<>"" then %>
						                <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPYME)>0,SUMA_S4TPYME,1),0)%>%</b></font>
						                </td>
						                <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
            						    
						                <td bgcolor="#BEE8FB" align="right" width="5%">
							                <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S3TPYME)>0,SUMA_S3TPYME,1),0)%>%</b></font>
						                </td>
						                <td bgcolor="#FFFFFF"  align="center" width="2%">
								            <font size=1 color=#483d8b face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>

							        <%else %>
							            <%if f4<>"" then%>
						                    <td bgcolor="#BEE8FB" align="right" width="5%">
							                    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*3/iif((SUMA_S4TPYME)>0,SUMA_S4TPYME,1),0)%>%</b></font>
						                    </td>
						                    <td bgcolor="#FFFFFF" align="center" width="2%">
								                <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							                </td>

            							    
							            <%end if %>	
        							    
							        <%end if %>	
    							    
							    <%end if %>	
							    
							<%end if %>	
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_META*12/iif((SUMA_SALIDASSINVENTAAAAPYME)>0,SUMA_SALIDASSINVENTAAAAPYME,1),0)%>%</b></font>
						    </td>
						  
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber((SUMA_S4TPYME)*100/iif((SUMA_StockCierrePYME4+SUMA_IT4PYME)>0,SUMA_StockCierrePYME4+SUMA_IT4PYME,1),2)%>%</b></font>
						    </td>
						    
						    <td bgcolor="#BEE8FB" align="right" width="5%">
							    <font size=1 color=#483d8b face=Arial><b><%=FormatNumber((SUMA_S3TPYME)*100/iif((SUMA_StockCierrePYME3+SUMA_IT3PYME)>0,SUMA_StockCierrePYME3+SUMA_IT3PYME,1),2)%>%</b></font>
						    </td>
	                    </tr>
                        <%

                        Do While not RS.EOF
                            if bgcolor="#FFFFFF"then
                                bgcolor="#F5F5F5"
                            else
                                bgcolor="#FFFFFF"
                            end if
                        %>
                        <tr>
	                        <td bgcolor="<%=bgcolor%>" rowspan="2">
		                        <a href="javascript:modificar('<%=trim(RS.Fields("Estudio"))%>');"><font size=1 color=#483d8b face=Arial><%=RS.Fields("Estudio")%></font></a>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" width="10%">
		                        <font size=1 color=#483d8b face=Arial>PARTICULARES</font>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("NUMCLIENTESPART"),0)%></font>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("NUMCASOSPART"),0)%></font>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("PROVCONSTPART")/1000,0)%></font>
	                        </td>	
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("StockCierrePART")/1000,0)%></font>
	                        </td>	
	                        <td bgcolor="#FFFFFF" align="center" width="1%">
								<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							</td>			
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasAAAAPART")/1000,0)%></font>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasAAAAConVentaPART")/1000,0)%></font>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SALIDASPART")/1000,0)%></font>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasConVentaPART")/1000,0)%></font>
	                        </td>
	                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial>0</font>
	                        </td>	
	                        
							<%if f1<>"" then %>
								<td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial>0%</font>
		                        </td>
		                        <td bgcolor="#FFFFFF" align="center" width="2%">
									<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
								</td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial>0%</font>
		                        </td>
		                        <td bgcolor="#FFFFFF" align="center" width="2%">
						            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
					            </td>
					            <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial>0%</font>
		                        </td>
		                        <td bgcolor="#FFFFFF" align="center" width="2%">
									<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
								</td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial>0%</font>
		                        </td>
		                        <td bgcolor="#FFFFFF" align="center" width="2%">
						            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
					            </td>
							<%else %>
								<%if f2<>"" then %>
		                            <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                            <font size=1 color=#483d8b face=Arial>0%</font>
		                            </td>
		                            <td bgcolor="#FFFFFF" align="center" width="2%">
						                <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
					                </td>
					                <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                            <font size=1 color=#483d8b face=Arial>0%</font>
		                            </td>
		                            <td bgcolor="#FFFFFF" align="center" width="2%">
									    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
								    </td>
		                            <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                            <font size=1 color=#483d8b face=Arial>0%</font>
		                            </td>
		                            <td bgcolor="#FFFFFF" align="center" width="2%">
						                <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
					                </td>
								<%else %>
									<%if f3<>"" then %>
										<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>0%</font>
				                        </td>
				                        <td bgcolor="#FFFFFF" align="center" width="2%">
											<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
										</td>
				                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>0%</font>
				                        </td>
				                        <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							            </td>
									<%else %>
										<%if f4<>"" then %>
											<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                            <font size=1 color=#483d8b face=Arial>0%</font>
				                            </td>
				                            <td bgcolor="#FFFFFF" align="center" width="2%">
								                <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
							                </td>
										<%end if %>

									<%end if %>

								<%end if %>

							<%end if %>


							<td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                       <font size=1 color=#483d8b face=Arial>0%</font>
	                        </td>

	                        <%if f1<>"" then %>
								<td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                            <font size=1 color=#483d8b face=Arial>
		                        	    <%= FORMATNUMBER(round(RS.Fields("S4TPART"),2)/(round((RS.Fields("StockCierrePART4")/100),2)+round((RS.Fields("IT4PART")/100),2)+1),2) %>%  	
	                            </td>
                                <td bgcolor="<%=bgcolor%>" align="right" width="5%">
	                                <font size=1 color=#483d8b face=Arial>
	                        	        <%= FORMATNUMBER(round(RS.Fields("S3TPART"),2)/(round((RS.Fields("StockCierrePART3")/100),2)+round((RS.Fields("IT3PART")/100),2)+1),2) %>%  	
                                </td>
                                <td bgcolor="<%=bgcolor%>" align="right" width="5%">
	                                <font size=1 color=#483d8b face=Arial>
	                        	        <%= FORMATNUMBER(round(RS.Fields("S2TPART"),2)/(round((RS.Fields("StockCierrePART2")/100),2)+round((RS.Fields("IT2PART")/100),2)+1),2) %>%  	
                                </td>
                                <td bgcolor="<%=bgcolor%>" align="right" width="5%">
	                                <font size=1 color=#483d8b face=Arial>
	                        	        <%= FORMATNUMBER(round(RS.Fields("S1TPART"),2)/(round((RS.Fields("StockCierrePART1")/100),2)+round((RS.Fields("IT1PART")/100),2)+1),2) %>%  	
                                </td>
							<%else %>
								<%if f2<>"" then %>
									<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>
					                        	<%= FORMATNUMBER(round(RS.Fields("S4TPART"),2)/(round((RS.Fields("StockCierrePART4")/100),2)+round((RS.Fields("IT4PART")/100),2)+1),2) %>%  	
				                        </td>
			                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
				                        <font size=1 color=#483d8b face=Arial>
				                        	<%= FORMATNUMBER(round(RS.Fields("S3TPART"),2)/(round((RS.Fields("StockCierrePART3")/100),2)+round((RS.Fields("IT3PART")/100),2)+1),2) %>%  	
			                        </td>
			                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
				                        <font size=1 color=#483d8b face=Arial>
				                        	<%= FORMATNUMBER(round(RS.Fields("S2TPART"),2)/(round((RS.Fields("StockCierrePART2")/100),2)+round((RS.Fields("IT2PART")/100),2)+1),2) %>%  	
			                        </td>
								<%else %>
									<%if f3<>"" then %>
										<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>
					                        	<%= FORMATNUMBER(round(RS.Fields("S4TPART"),2)/(round((RS.Fields("StockCierrePART4")/100),2)+round((RS.Fields("IT4PART")/100),2)+1),2) %>%  	
				                        </td>
				                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>
					                        	<%= FORMATNUMBER(round(RS.Fields("S3TPART"),2)/(round((RS.Fields("StockCierrePART3")/100),2)+round((RS.Fields("IT3PART")/100),2)+1),2) %>%  	
				                        </td>
									
									<%else %>
										<%if f4<>"" then %>
											<td bgcolor="#007DC5" align="center" width="5%">
												<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Marzo</b></font>
											</td>
										<%end if %>

									<%end if %>

								<%end if %>

							<%end if %>
	                        	                        						                        	                        
	                        		                        						                        
						</tr>	
						<tr>
							<td bgcolor="<%=bgcolor%>" width="10%">
			                        <font size=1 color=#483d8b face=Arial>PYME</font>
		                        </td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("NUMCLIENTESPYME"),0)%></font>
		                        </td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("NUMCASOSPYME"),0)%></font>
		                        </td>	
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("PROVCONSTPYME")/1000,0)%></font>
		                        </td>	
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("StockCierrePYME")/1000,0)%></font>
		                        </td>	
		                        <td bgcolor="#FFFFFF" align="center" width="1%">
									<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
								</td>			
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasAAAAPYME")/1000,0)%></font>
		                        </td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasAAAAConVentaPYME")/1000,0)%></font>
		                        </td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SALIDASPYME")/1000,0)%></font>
		                        </td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("SalidasConVentaPYME")/1000,0)%></font>
		                        </td>
		                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
		                        <font size=1 color=#483d8b face=Arial>0</font>
	                        </td>	
								<%if f1<>"" then %>
									<td bgcolor="<%=bgcolor%>" align="right" width="5%">
				                        <font size=1 color=#483d8b face=Arial>0%</font>
			                        </td>
			                        <td bgcolor="#FFFFFF" align="center" width="2%">
										<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
									</td>
			                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
				                        <font size=1 color=#483d8b face=Arial>0%</font>
			                        </td>
			                        <td bgcolor="#FFFFFF" align="center" width="2%">
										<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
									</td>
									<td bgcolor="<%=bgcolor%>" align="right" width="5%">
				                        <font size=1 color=#483d8b face=Arial>0%</font>
			                        </td>
			                        <td bgcolor="#FFFFFF" align="center" width="2%">
										<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
									</td>
									<td bgcolor="<%=bgcolor%>" align="right" width="5%">
				                        <font size=1 color=#483d8b face=Arial>0%</font>
			                        </td>
			                        <td bgcolor="#FFFFFF" align="center" width="2%">
										<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
									</td>
								<%else %>
									<%if f2<>"" then %>
										<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>0%</font>
				                        </td>
				                        <td bgcolor="#FFFFFF" align="center" width="2%">
											<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
										</td>
				                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>0%</font>
				                        </td>
				                        <td bgcolor="#FFFFFF" align="center" width="2%">
											<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
										</td>
										<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>0%</font>
				                        </td>
				                        <td bgcolor="#FFFFFF" align="center" width="2%">
											<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
										</td>
									<%else %>
										<%if f3<>"" then %>
											<td bgcolor="<%=bgcolor%>" align="right" width="5%">
						                        <font size=1 color=#483d8b face=Arial>0%</font>
					                        </td>
					                        <td bgcolor="#FFFFFF" align="center" width="2%">
												<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
											</td>
					                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
						                        <font size=1 color=#483d8b face=Arial>0%</font>
					                        </td>
					                        <td bgcolor="#FFFFFF" align="center" width="2%">
												<font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
											</td>
										<%else %>
											<%if f4<>"" then %>
												<td bgcolor="<%=bgcolor%>" align="right" width="5%">
						                        <font size=1 color=#483d8b face=Arial>0%</font>
					                            </td>
					                            <td bgcolor="#FFFFFF" align="center" width="2%">
												    <font size=1 color=#FFFFFF face=Arial><b><img src="imagenes/sem_rojo.png" border=0 alt="semtemporal" title="Buscar" width="12" border=0></b></font>
											    </td>
					                       
											<%end if %>

										<%end if %>

									<%end if %>

								<%end if %>

								<td bgcolor="<%=bgcolor%>" align="right" width="5%">
			                        <font size=1 color=#483d8b face=Arial>0%</font>
		                        </td>

		                        <%if f1<>"" then %>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Diciembre</b></font>
									</td>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Setiembre</b></font>
									</td>	
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Junio</b></font>
									</td>
									<td bgcolor="#007DC5" align="center" width="5%">
										<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Marzo</b></font>
									</td>
								<%else %>
									<%if f2<>"" then %>
										<td bgcolor="#007DC5" align="center" width="5%">
											<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Setiembre</b></font>
										</td>	
										<td bgcolor="#007DC5" align="center" width="5%">
											<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Junio</b></font>
										</td>
										<td bgcolor="#007DC5" align="center" width="5%">
											<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Marzo</b></font>
										</td>
									<%else %>
										<%if f3<>"" then %>
											<td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>
					                        	<%= FORMATNUMBER(round(RS.Fields("S4TPYME"),2)/(round((RS.Fields("StockCierrePYME4")/100),2)+round((RS.Fields("IT4PYME")/100),2)+1),2) %>%  	
				                        </td>
				                        <td bgcolor="<%=bgcolor%>" align="right" width="5%">
					                        <font size=1 color=#483d8b face=Arial>
					                        	<%= FORMATNUMBER(round(RS.Fields("S3TPYME"),2)/(round((RS.Fields("StockCierrePYME3")/100),2)+round((RS.Fields("IT3PYME")/100),2)+1),2) %>%  	
				                        </td>
											
										<%else %>
											<%if f4<>"" then %>
												<td bgcolor="#007DC5" align="center" width="5%">
													<font size=1 color=#FFFFFF face=Arial><b>% Alcance Meta Marzo</b></font>
												</td>
											<%end if %>

										<%end if %>

									<%end if %>

								<%end if %>
	                        
		                        	                        						                        	                        
		                        		                        						                        
							</tr>	


						                            
                        <%


                        
                        
                        RS.MoveNext
                        Loop
                        RS.Close 					
					    %>	

					  <tr>
						    <td bgcolor="#007DC5">
							    <font size=1 color=#FFFFFF face=Arial><b>TOTAL</b></font>
						    </td>
						    <td bgcolor="#007DC5">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	
						    <td bgcolor="#FFFFFF" align="center" width="1%">
								<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							</td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>	
						    <%if f1<>"" then%>
						        <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>
							    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>
							    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>
						        <td bgcolor="#007DC5" align="right" width="5%">
							        <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						        </td>
						        <td bgcolor="#FFFFFF" align="center" width="2%">
								    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							    </td>
						    <%else %>
						        <%if f2<>"" then%>

							        <td bgcolor="#007DC5" align="right" width="5%">
							        <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						            </td>
						            <td bgcolor="#FFFFFF" align="center" width="2%">
								        <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							        </td>
							        <td bgcolor="#007DC5" align="right" width="5%">
							        <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						            </td>
						            <td bgcolor="#FFFFFF" align="center" width="2%">
								        <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							        </td>
						            <td bgcolor="#007DC5" align="right" width="5%">
							            <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						            </td>
						            <td bgcolor="#FFFFFF" align="center" width="2%">
								        <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							        </td>
						        <%else %>
						            <%if f3<>"" then%>

							            <td bgcolor="#007DC5" align="right" width="5%">
							            <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						                </td>
						                <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							            </td>
						                <td bgcolor="#007DC5" align="right" width="5%">
							                <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						                </td>
						                <td bgcolor="#FFFFFF" align="center" width="2%">
								            <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							            </td>
						            <%else %>
    						            <%if f4<>"" then%>

							                <td bgcolor="#007DC5" align="right" width="5%">
							                <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						                    </td>
						                    <td bgcolor="#FFFFFF" align="center" width="2%">
								                <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
							                </td>
        						        
						                <%end if %>
						            <%end if %>
						        <%end if %>
						    <%end if %>
							
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						    </td>
						    <td bgcolor="#007DC5" align="right" width="5%">
							    <font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
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
		
		<!--cargando--><script language=javascript>		                   document.getElementById("imgloading").style.display = "none";</script>	
		<%	
		''end if
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
    //window.open("index.html","_top");
    window.open("index.html", "sistema");
    window.close();
</script>
<%
end if
%>



