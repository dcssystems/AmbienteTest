<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repsaeestudios.asp") then
	
	    codestudio=obtener("codestudio")
	    plaza=obtener("plaza")
	    segmento=obtener("segmento")
	    confase=obtener("confase")
	    
        filtrobuscador = ""
		if codestudio<>"" then
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
			    filtrobuscador = filtrobuscador & " where A.faseactual is not null"
			else
			    filtrobuscador = filtrobuscador & " and A.faseactual is not null"
			end if
		end if					
			    
		sql="select A.Plaza, " & _
		    "A.CODDATO, " & _
		            "IsNull(A.nomb_estu,'NO IDENTIFICADO') as ESTUDIO, " & _
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
	            "from FB.CentroCobranzas.dbo.Vista_Detalle_Casos_JUD A " & _
	            "where A.FASIGNA<=DATEADD(m,-2,getdate()) " & _
	            "group by A.CONTRATO,A.Plaza,A.CODDATO,A.nomb_estu,SEGMENTO_RIESGO,A.JUD) A " & _
            "group by A.Plaza,A.CODDATO,IsNull(A.nomb_estu,'NO IDENTIFICADO') " & _
            "order by A.Plaza,MontoSolesPART desc,MontoSolesPYME desc"
            ''"order by A.Plaza,IsNull(A.nomb_estu,'NO IDENTIFICADO')"

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
		    SUMA_MontoSoles_PROV_PYME=SUMA_MontoSoles_PROV_PYME + RS.Fields("MontoSolesPART")
		    SUMA_Cantidad_PROV_PYME=SUMA_Cantidad_PROV_PYME + RS.Fields("CasosPART")
		    SUMA_MontoSoles_PROV_PYME_CF=SUMA_MontoSoles_PROV_PYME_CF + RS.Fields("MontoSolesConFasePART")
		    SUMA_Cantidad_PROV_PYME_CF=SUMA_Cantidad_PROV_PYME_CF + RS.Fields("CasosConFasePART")
        RS.MoveNeXt
        Loop    

		SUMA_MontoSoles_PART=SUMA_MontoSoles_LIMA_PART + SUMA_MontoSoles_PROV_PART
		SUMA_Cantidad_PART=SUMA_Cantidad_LIMA_PART + SUMA_Cantidad_LIMA_PYME 
		SUMA_MontoSoles_PYME=SUMA_MontoSoles_LIMA_PYME + SUMA_MontoSoles_PROV_PYME
		SUMA_Cantidad_PYME=SUMA_Cantidad_PROV_PYME + SUMA_Cantidad_PROV_PYME
		SUMA_MontoSoles_PART_CF=SUMA_MontoSoles_LIMA_PART_CF + SUMA_MontoSoles_PROV_PART_CF
		SUMA_Cantidad_PART_CF=SUMA_Cantidad_LIMA_PART_CF + SUMA_Cantidad_LIMA_PYME_CF
		SUMA_MontoSoles_PYME_CF=SUMA_MontoSoles_LIMA_PYME_CF + SUMA_MontoSoles_PROV_PYME_CF
		SUMA_Cantidad_PYME_CF=SUMA_Cantidad_PROV_PYME_CF + SUMA_Cantidad_PROV_PYME_CF
		%>
		
		
		<!--Ojo esta ventana siempre es flotante-->
		<html>
			<title>SAE - Ranking por Estudios</title>
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
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF" onload="cuadratabla();" onresize="cuadratabla();">
			<form name=formula method=post>
					<table border=0 cellspacing=2 cellpadding=4 width=100%>
					<tr>	
						<td bgcolor="#F5F5F5" align="center">			
							<font size=4 color=#483d8b face=Arial><b>SAE - Ranking por Estudio</b></font>
						</td>
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
							<font size=1 color=#FFFFFF face=Arial><b>S/.MM</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>S/.MM</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" width=15>
							<font size=1 color=#FFFFFF face=Arial><b>&nbsp;</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>S/.MM</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>N° Casos</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>%</b></font>
						</td>
						<td bgcolor="#007DC5" align="center" width=5%>
							<font size=1 color=#FFFFFF face=Arial><b>S/.MM</b></font>
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
							<font size=1 color=#483d8b face=Arial><b>Total</b></font>
						</td>
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b>100.00%</b></font>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_CF/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PART_CF%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</b></font>
						</td>	
						<td bgcolor="#BEE8FB" align="center" width=15><img src="imagenes/<%=iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PART_CF*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b>100.00%</b></font>
						</td>		
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_CF/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PYME_CF%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right" width=5%>
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PYME_CF*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</b></font>
						</td>
						<td bgcolor="#BEE8FB" align="center" width=15><img src="imagenes/<%=iif(SUMA_MontoSoles_PYME_CF*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PYME_CF*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
					</tr>					
					<!--LIMA-->
					<tr>
						<td bgcolor="#BEE8FB">
							<font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Lima</b></font>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PART%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</b></font>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART_CF/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PART_CF%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1),2)%>%</b></font>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_PART_CF*100/iif(SUMA_MontoSoles_LIMA_PART>0,SUMA_MontoSoles_LIMA_PART,1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PYME%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</b></font>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME_CF/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_LIMA_PYME_CF%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1),2)%>%</b></font>
						</td>
						<td bgcolor="#BEE8FB" align="center"><img src="imagenes/<%=iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=90,"sem_verde",iif(SUMA_MontoSoles_LIMA_PYME_CF*100/iif(SUMA_MontoSoles_LIMA_PYME>0,SUMA_MontoSoles_LIMA_PYME,1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
					</tr>	
					<%
                        RS.Filter=""
                        RS.Filter=" Plaza='LIMA' "
                        Do While not RS.EOF
                        %>
                        <tr>
	                        <td bgcolor="#F5F5F5">
		                        <font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESTUDIO")%></font>
	                        </td>
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPART")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</font>
	                        </td>		
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePART")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1),2)%>%</font>
	                        </td>	
	                        <td bgcolor="#F5F5F5" align="center"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPYME")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</font>
	                        </td>		
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePYME")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1),2)%>%</font>
	                        </td>
	                        <td bgcolor="#F5F5F5" align="center"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
                        </tr>	                        
                        <%
                        RS.MoveNext
                        Loop 					
					%>			
					<!--PROVINCIA-->	
					<tr>
						<td bgcolor="#BEE8FB">
							<font size=1 color=#483d8b face=Arial><b>&nbsp;&nbsp;Provincia</b></font>
						</td>
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PART%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</b></font>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART_CF/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PART_CF%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1),2)%>%</b></font>
						</td>	
						<td bgcolor="#BEE8FB" align="center"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_PART_CF*100/iif(SUMA_MontoSoles_PROV_PART>0,SUMA_MontoSoles_PROV_PART,1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PYME%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</b></font>
						</td>		
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME_CF/1000000,0)%></b></font>
						</td>						
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=SUMA_Cantidad_PROV_PYME_CF%></b></font>
						</td>				
						<td bgcolor="#BEE8FB" align="right">
							<font size=1 color=#483d8b face=Arial><b><%=FormatNumber(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1),2)%>%</b></font>
						</td>
						<td bgcolor="#BEE8FB" align="center"><img src="imagenes/<%=iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=90,"sem_verde",iif(SUMA_MontoSoles_PROV_PYME_CF*100/iif(SUMA_MontoSoles_PROV_PYME>0,SUMA_MontoSoles_PROV_PYME,1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
					</tr>	
                        <%
                        RS.Filter=""
                        RS.Filter=" Plaza='PROVINCIA' "
                        Do While not RS.EOF
                        %>
                        <tr>
	                        <td bgcolor="#F5F5F5">
		                        <font size=1 color=#483d8b face=Arial>&nbsp;&nbsp;&nbsp;&nbsp;<%=RS.Fields("ESTUDIO")%></font>
	                        </td>
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPART")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPART")*100/iif(SUMA_MontoSoles_PART>0,SUMA_MontoSoles_PART,1),2)%>%</font>
	                        </td>		
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePART")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1),2)%>%</font>
	                        </td>	
	                        <td bgcolor="#F5F5F5" align="center"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosPYME")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesPYME")*100/iif(SUMA_MontoSoles_PYME>0,SUMA_MontoSoles_PYME,1),2)%>%</font>
	                        </td>		
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")/1000000,0)%></font>
	                        </td>						
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=RS.Fields("CasosConFasePYME")%></font>
	                        </td>				
	                        <td bgcolor="#F5F5F5" align="right">
		                        <font size=1 color=#483d8b face=Arial><%=FormatNumber(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1),2)%>%</font>
	                        </td>
	                        <td bgcolor="#F5F5F5" align="center"><img src="imagenes/<%=iif(RS.Fields("MontoSolesConFasePYME")*100/iif(RS.Fields("MontoSolesPYME")>0,RS.Fields("MontoSolesPYME"),1)>=90,"sem_verde",iif(RS.Fields("MontoSolesConFasePART")*100/iif(RS.Fields("MontoSolesPART")>0,RS.Fields("MontoSolesPART"),1)>=80,"sem_amarillo","sem_rojo"))%>.png" width="12" border=0></td>
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
			</form>							
			</body>
		</html>	
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

