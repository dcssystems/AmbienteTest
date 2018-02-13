<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admfacultad.asp") then
		buscador=obtener("buscador")
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
	%>
		<html>
		<!--cargando--><img src="imagenes/loading.gif" border="0" id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
			<link rel="stylesheet" href="assets/css/css/animation.css"/>
			<link rel="stylesheet" href="assets/css/custom.css" />
			<link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"/>
			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->	
			
			
	
    <script>
      function toggleCodes(on) {
        var obj = document.getElementById('icons');
      
        if (on) {
          obj.className += ' codesOn';
        } else {
          obj.className = obj.className.replace(' codesOn', '');
        }
      }
      
    </script>


    	
		<script language="javascript" src="scripts/TablaDinamica.js"></script>
		<script language="javascript">
		var ventanafacultad;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanafacultad=global_popup_IWTSystem(ventanafacultad,"dcs_nuevofacultad.asp?vistapadre=" + window.name + "&paginapadre=dcs_admfacultad.asp&codfacultad=" + codigo,"NewFacultad","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 180)/2 - 30) + ",height=180,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregar()
		{
			ventanafacultad=global_popup_IWTSystem(ventanafacultad,"dcs_nuevofacultad.asp?vistapadre=" + window.name + "&paginapadre=dcs_admfacultad.asp","NewFacultad","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 180)/2 - 30) + ",height=180,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
		
		</head>
		
		<script language="javascript">
			<%
				idcampana=2 ''obtener("idcampaña")
				sql="select GlosaCampo,ROW_NUMBER () over (order by NroCampo) as Orden,CampoCalculado,Formula,Condicion,IDCampañaCampo,TipoCampo,FlagNroDocumento,anchocolumna,aligncabecera,aligndetalle,alignpie,decimalesnumero,formatofecha " & chr(10) & _
                    "from Campaña_Campo " & chr(10) & _
                    "where IDTipoCampaña in (select IDTipoCampaña from Campaña where idcampaña=2) " & chr(10) & _
                    "and Nivel=1 and Visible=1 " & chr(10) & _
                    "order by Orden"
				consultar sql,RS3	
				nrocampos=RS3.RecordCount
				glosacampos=""
				glosavisible=""
				glosaancho=""
				glosaalineamiento=""
				Do while not RS3.EOF 
					glosacampos=glosacampos & ",'" & RS3.Fields("GlosaCampo") & "'"
					glosavisible=glosavisible & ",'true'"
					glosaancho=glosaancho & ",'" & RS3.Fields("anchocolumna") & "'"
					glosaaligncabecera=glosaaligncabecera & ",'" & RS3.Fields("aligncabecera") & "'"
					glosaaligndetalle=glosaaligndetalle & ",'" & RS3.Fields("aligndetalle") & "'"
					glosaalignpie=glosaalignpie & ",'" & RS3.Fields("alignpie") & "'"
					glosadecimalesnumero=glosadecimalesnumero & ",'" & RS3.Fields("decimalesnumero") & "'"
					glosaformatofecha=glosaformatofecha & ",'" & RS3.Fields("formatofecha") & "'"
					glosapie=glosapie & ",'&nbsp;'"
					glosapiefunciones=glosapiefunciones & ",''"
				RS3.MoveNext 
				Loop
				RS3.MoveFirst
			%>
			rutaimgcab="imagenes/"; 
		  //Configuración general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=<%=nrocampos + 1%>;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('IDCampanaPersona'<%=glosacampos%>);
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(false<%=glosavisible%>);
		    anchocolumna[tabla] =  new Array( ''<%=glosaancho%>);
		    aligncabecera[tabla] = new Array('left'<%=glosaaligncabecera%>);
		    aligndetalle[tabla] = new Array('left'<%=glosaaligndetalle%>);
		    alignpie[tabla] =     new Array('left'<%=glosaalignpie%>);
		    decimalesnumero[tabla] = new Array(-1<%=glosadecimalesnumero%>);
		    formatofecha[tabla] =   new Array(''<%=glosaformatofecha%>);


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=idcampanapersona-id- value=-c0->' + '<a href="javascript:modificar(-id-);">-valor-</a>';
				<%
				indicecampo=0
				Do while not RS3.EOF 
				    indicecampo=indicecampo + 1
				    %>objetofomulario[tabla][<%=indicecampo%>]='<a href="javascript:modificar(-id-);">-valor-</a>';
				    <%
				RS3.MoveNext 
				Loop
				RS3.MoveFirst				
				%>					
					
		    filtrardatos[tabla]=0; //define si carga auto el filtro
		    filtrofomulario[tabla] = new Array();
		    tipofiltrofomulario[tabla] = new Array();
		    	filtrofomulario[tabla][0]='';
                <%
				indicecampo=0
				Do while not RS3.EOF 
				    indicecampo=indicecampo + 1
				    %>filtrofomulario[tabla][<%=indicecampo%>]='';
				    <%
				RS3.MoveNext 
				Loop
				RS3.MoveFirst				
				%>	
				
				
		    valorfiltrofomulario[tabla] = new Array();
				valorfiltrofomulario[tabla][0]='';
                <%
				indicecampo=0
				Do while not RS3.EOF 
				    indicecampo=indicecampo + 1
				    %>valorfiltrofomulario[tabla][<%=indicecampo%>]='';
				    <%
				RS3.MoveNext 
				Loop
				RS3.MoveFirst				
				%>	

		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		if buscador<>"" then
			filtrobuscador = " where (f.codfacultad like '%" & buscador & "%' or f.descripcion like '%" & buscador & "%' or g.descripcion like '%" & buscador & "%' or f.orden like '%" & buscador & "%') "
		end if
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if		
		
		contadortotal=0
		sql="select count(*) from facultad f inner join grupofacultad g on f.codgrupofacultad = g.codgrupofacultad " & filtrobuscador 
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

        sql="select A.IDCampañaPersona,A.NroDocumento,C.IDCampañaCampo,C.NroCampo,C.TipoCampo,D.ValorTexto,D.ValorEntero,D.ValorFloat,D.ValorFecha " & chr(10) & _
            "from Campaña_Persona A " & chr(10) & _
            "inner join Campaña B " & chr(10) & _
            "on A.IDCampaña=B.IDCampaña " & chr(10) & _
            "inner join Campaña_Campo C " & chr(10) & _
            "on B.IDTipoCampaña=C.IDTipoCampaña and C.Nivel=1 and C.Visible=1 " & chr(10) & _
            "inner join Campaña_Detalle D " & chr(10) & _
            "on A.IDCampañaPersona=D.IDCampañaPersona and C.IDCampañaCampo=D.IDCampañaCampo " & chr(10) & _
            "where A.IDCampaña=" & IDCampana 
		consultar sql,RS2            
		
		if pag>1 then					
		sql="select top 0 *,f.descripcion as descrip,g.descripcion as grupo, f.orden from facultad f inner join grupofacultad g on f.codgrupofacultad = g.codgrupofacultad where " & filtrobuscador1 & " f.codfacultad not in (select top " & topnovisible & " f.codfacultad from facultad f inner join grupofacultad g on f.codgrupofacultad = g.codgrupofacultad " & filtrobuscador & " order by f.codfacultad) order by f.codfacultad" 
		else
		sql="select A.IDCampañaPersona,A.NroDocumento from Campaña_Persona A where A.idcampaña=2 order by A.IDCampañaPersona"
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF 
							
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]=<%=RS.Fields("IDCampañaPersona")%>;
                <%
				indicecampo=0
				Do while not RS3.EOF
				    indicecampo=indicecampo + 1
   				    if RS3.Fields("FlagNroDocumento")=0 then
				        RS2.Filter=" IDCampañaPersona=" & RS.Fields("IDCampañaPersona") & " and IDCampañaCampo=" & RS3.Fields("IDCampañaCampo") & " "
    				    
				        Select Case RS3.Fields("TipoCampo")
				        case 1 
				                Response.Write "datos[tabla][" & contador & "][" & indicecampo & "]='" & RS2.Fields("ValorTexto") & "';" & chr(10)
				        case 2 
				                Response.Write "datos[tabla][" & contador & "][" & indicecampo & "]=" & RS2.Fields("ValorEntero") & ";" & chr(10)
				        case 3 
				                Response.Write "datos[tabla][" & contador & "][" & indicecampo & "]=" & RS2.Fields("ValorFloat") & ";" & chr(10)
				        case 4
				                if not IsNull(RS2.Fields("ValorFecha")) then
				                    valorfecha="new Date(" & Year(RS2.Fields("ValorFecha")) & "," & Month(RS2.Fields("ValorFecha"))-1 & "," & Day(RS2.Fields("ValorFecha")) & "," & Hour(RS2.Fields("ValorFecha")) & "," & Minute(RS2.Fields("ValorFecha")) & "," & Second(RS2.Fields("ValorFecha")) & ")"
				                else
				                    valorfecha="null"
				                end if
				                Response.Write "datos[tabla][" & contador & "][" & indicecampo & "]=" & valorfecha & ";" & chr(10)
				        End Select
				    else
				        Response.Write "datos[tabla][" & contador & "][" & indicecampo & "]='" & RS.Fields("NroDocumento") & "';" & chr(10)
				    end if
				RS3.MoveNext 
				Loop
				RS3.MoveFirst				

			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;'<%=glosapie%>);
		    piefunciones[tabla] = new Array(''<%=glosapiefunciones%>); 


		    //Se escriben las opciones para los selects que contenga
		    posicionselect[tabla]=new Array();
		    nombreselect[tabla]=new Array();
		    opcionesvalor[tabla]=new Array();
		    opcionestexto[tabla]=new Array();
		    //Finaliza configuracion de tabla 0

				
			    
		    funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
		    funcionagrega[tabla]='agregar();';

		</script> 
		
		<%if contador=0 then%>
		
		<body topmargin="0" leftmargin="0">
			<form name="formula" method="post">
				<table width="100%" cellpadding="4" cellspacing="0">	
					<tr class="fondo-orange">
						<td class="text-orange"><font size="2" face="Raleway"><b>Facultad (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-doc">&#xe808;</i></a></td>
						<td class="text-orange" align="middle" width="250"><font size="2" face="Raleway">Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size="20" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
						<td class="text-orange" align="left"><a href="javascript:buscar();"><i class="demo-icon icon-search">&#xe80c;</i></a></td>
					</tr>
				</table>
		<%else		
		%>
		<body topmargin="0" leftmargin="0"><!--onload="inicio();"-->
			<form name="formula" method="post">
				<table width="100%" cellpadding="4" cellspacing="0" border="0">		
					<tr class="fondo-orange">
						<td class="text-orange" align="left"><font size="2" face="Raleway"><b>Facultad (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-doc">&#xe808;</i></a>&nbsp;&nbsp;<a href="javascript:exportar();"><i class="demo-icon icon-file-excel">&#xf1c3;</i></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align="middle"></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><i class="demo-icon icon-download">&#xe814;</i></a><%end if%></b></font></td>
						<!--<td bgcolor="#F5F5F5" align="left"><font size="2" face="Raleway" color=#00529B><b>Grupo Facultad (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-doc">&#xe808;</i></a><!--&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align="middle"></a>&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align="middle"></a><%if expimp="1" then%>&nbsp;&nbsp;<a href='exportados/<%=nombrearchivo%>.xls','VerExport'><i class="demo-icon icon-download">&#xe814;</i></a><%end if%></b></font></td>-->
						<td class="text-orange" align="middle" width="250"><font size="2" face="Raleway">Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size="20" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
						<td class="text-orange" align="left"><a href="javascript:buscar();"><i class="demo-icon icon-search">&#xe80c;</i></a></td>
						<td class="text-orange" align="right" width="180"><font size="2" face="Raleway">Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
					</tr>	
				</table>
				<div id="tabla0"> 
				</div>
		<%end if%>
		<input type="hidden" name="actualizarlista" value="">
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="pag" value="<%=pag%>">	
		</form>
		<script language="javascript">
			inicio();
		</script>					
		</body>
		<!--cargando--><script language="javascript">document.getElementById("imgloading").style.display="none";</script>		
		</html>
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
					consulta_exp="select 'Cod.Facultad','Grupo','Descripcion','Pagina','Orden'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select f.codfacultad,g.descripcion,f.descripcion,f.pagina,f.orden " & _
								 "from CobranzaCM.dbo.facultad f inner join CobranzaCM.dbo.grupofacultad g on f.codgrupofacultad = g.codgrupofacultad " & filtrobuscador & " order by f.codfacultad" 
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
		window.open("dcs_userexpira.asp","_top");
	</script>
	<%	
	end if
	RS2.Close
	RS3.Close
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



