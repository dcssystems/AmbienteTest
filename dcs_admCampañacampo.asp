<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admCampa�acampo.asp") then
		buscador=obtener("buscador")
		''Codigo exp excel - se repite
		expimp=obtener("expimp")
		if expimp="1" then
			sql="SELECT descripcion,valortexto1 FROM parametro WHERE descripcion='RutaFisicaExportar' OR descripcion='RutaWebExportar'"
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
		<link rel="stylesheet" href="assets/css/css/animation.css" />
		<link rel="stylesheet" href="assets/css/custom.css" />
		
		<link href="https://fonts.googleapis.com/css?family=Fira+Sans+Condensed:400|Raleway:500&amp;subset=greek,greek-ext,latin-ext" rel="stylesheet"> 
		
		
		<!--link href="https://fonts.googleapis.com/css?family=Raleway&amp;subset=latin-ext" rel="stylesheet"-->
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
			ventanafacultad=global_popup_IWTSystem(ventanafacultad,"dcs_nuevoCliente.asp?vistapadre=" + window.name + "&paginapadre=dcs_admCliente.asp&IDCliente=" + codigo,"NewCliente","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 220)/2 - 30) + ",height=180,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregar()
		{
			ventanafacultad=global_popup_IWTSystem(ventanafacultad,"dcs_nuevoCliente.asp?vistapadre=" + window.name + "&paginapadre=dcs_admCliente.asp","NewCliente","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 180)/2 - 30) + ",height=180,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
			rutaimgcab="imagenes/"; 
		  //Configuraci�n general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=8;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('IDCampa�aCampo','Tipo Campa�a','GlosaCampo','Nivel','Tipo Campo','FlagNroDocumento','Visible','Editar');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(true, true, true ,true,true, true,true, true);
		    anchocolumna[tabla] =  new Array( '6%','20%', '6%' , '30%','6%' ,'5%','4%' ,'');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left','left','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1,-1,-1 ,-1 , -1 ,-1 , -1 );
		    formatofecha[tabla] =   new Array(''  ,'','','' ,'','','','' );


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=IDCampa�aCampo-id- value=-c0->' + '<a href="javascript:modificar(-id-);">-valor-</a>';
				objetofomulario[tabla][1]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][2]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][3]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][4]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][5]='<a href=javascript:modificar("-id-");>-valor-</a>';
				objetofomulario[tabla][6]=objetodatos("checkbox",tabla,"Visible","","","");
				objetofomulario[tabla][7]='<a href="javascript:modificar(-id-);"><i class="demo-icon2 icon-pencil-squared">&#xf14b;</i></a>';
				
										
					
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
				
										
					
		    valorfiltrofomulario[tabla] = new Array();
				valorfiltrofomulario[tabla][0]='';
				valorfiltrofomulario[tabla][1]='';
				valorfiltrofomulario[tabla][2]='';
				valorfiltrofomulario[tabla][3]='';
				valorfiltrofomulario[tabla][4]='';
				valorfiltrofomulario[tabla][5]='';
				valorfiltrofomulario[tabla][6]='';
				valorfiltrofomulario[tabla][7]='';


		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		if buscador<>"" then
			filtrobuscador = " WHERE (a.GlosaCampo LIKE '%" & buscador & "%' or b.Descripcion LIKE '%" & buscador & "%' ) "
		end if
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " AND "
		end if		
		
		contadortotal=0
		sql="SELECT COUNT(*) FROM Campa�a_Campo a inner join TipoCampa�a b on a.IDTipoCampa�a = b.IDTipoCampa�a " & filtrobuscador 
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
		sql="SELECT TOP " & cantidadxpagina & " a.IDCampa�aCampo,a.GlosaCampo,a.Nivel,a.TipoCampo,a.FlagNroDocumento,a.Visible, b.Descripcion as Descampa�a from Campa�a_Campo a inner join TipoCampa�a b on a.IDTipoCampa�a = b.idtipocampa�a where " & filtrobuscador1 & " a.IDCampa�aCampo NOT  IN (SELECT TOP " & topnovisible & " a.IDCampa�aCampo FROM Campa�a_Campo a inner join TipoCampa�a b on a.IDTipoCampa�a = b.idtipocampa�a  " & filtrobuscador & " ORDER BY a.IDCampa�aCampo) ORDER BY a.IDCampa�aCampo" 
		else
		sql="SELECT TOP " & cantidadxpagina & " a.IDCampa�aCampo,a.GlosaCampo,a.Nivel,a.TipoCampo,a.FlagNroDocumento,a.Visible, b.Descripcion as 'Descampa�a' from Campa�a_Campo a inner join TipoCampa�a b on a.IDTipoCampa�a = b.IDTipoCampa�a " & filtrobuscador & " ORDER BY a.IDCampa�aCampo" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF 
				if obtener("actualizarlista")<>"" and obtener("IDCampa�aCampo" & RS.Fields("IDCampa�aCampo"))<>"" then
					
					if obtener("Visible" & RS.Fields("IDCampa�aCampo"))<>"" then
						Visible="1"
					else
						Visible="0"
					end if		

		
						if 	int(Visible) <> rs.Fields("Visible") then
							sql="UPDATE Campa�a_Campo SET Visible=" & Visible & " WHERE IDCampa�aCampo=" & rs.Fields("IDCampa�aCampo") 
									'response.write "query:" & sql
							conn.Execute sql
						end if	
				end if 

										
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]='<%=RS.Fields("IDCampa�aCampo")%>';
				datos[tabla][<%=contador%>][1]='<%=rs.Fields("Descampa�a")%>';
				datos[tabla][<%=contador%>][2]='<%=rs.Fields("GlosaCampo")%>';
				datos[tabla][<%=contador%>][3]='<%=rs.Fields("Nivel")%>';
				datos[tabla][<%=contador%>][4]='<%=rs.Fields("TipoCampo")%>';
				datos[tabla][<%=contador%>][5]='<%=rs.Fields("FlagNroDocumento")%>';
				datos[tabla][<%=contador%>][6]=<%if obtener("actualizarlista")<>"" and obtener("IDCampa�aCampo" & RS.Fields("IDCampa�aCampo"))<>"" then%><%if int(Visible)=1 then%>'checked'<%else%>' '<%end if%><%else%><%if rs.Fields("Visible")=1 then%>'checked'<%else%>' '<%end if%><%end if%>;						
				datos[tabla][<%=contador%>][7]='';
							
		<%
			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','','','',''); 


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
			<body topmargin="0" leftmargin="0" face="Fira Sans Condensed">
				<form name="formula" method="post">
					<table width="100%" cellpadding="4" cellspacing="0">	
						<tr class="fondo-orange">
							<td class="text-orange"><font size="2" face="Raleway" ><b>Campa�a Campo (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border="0" alt="Nuevo" title="Nuevo" align="middle"></a></td>
							<td class="text-orange" align="middle" width="250"><font size="2" face="Raleway">Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size="20" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
							<td class="text-orange" align="left"><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border="0" alt="Buscar" title="Buscar" align="middle"></a></td>
						</tr>
				</table>			
		<%else		
		%>
			<body topmargin="0" leftmargin="0"><!--onload="inicio();"-->
				<form name="formula" method="post">
					<table width="100%" cellpadding="4" cellspacing="0" border="0">		
						<tr class="fondo-orange">
							<td class="text-orange" align="left"><font size="2" face="Raleway"><b>Campa�a Campo (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-doc">&#xe808;</i></a>&nbsp;&nbsp;<a href="javascript:exportar();"><i class="demo-icon icon-file-excel">&#xf1c3;</i></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><i class="demo-icon icon-download">&#xe814;</i></a><%end if%></b></font></td>
							<!--<td bgcolor="#F5F5F5" align=left><font size=2 face=Raleway color=#00529B><b>Grupo Facultad (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><i class="demo-icon icon-floppy">&#xe809;</i></a>&nbsp;&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-doc">&#xe808;</i></a><!--&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a>&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a><%if expimp="1" then%>&nbsp;&nbsp;<a href='exportados/<%=nombrearchivo%>.xls','VerExport'><i class="demo-icon icon-download">&#xe814;</i></a><%end if%></b></font></td>-->
							<td class="text-orange" align="middle" width="250"><font size="2" face="Raleway">Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size="20" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
							<td class="text-orange" align="left"><a href="javascript:buscar();"><i class="demo-icon icon-search">&#xe80c;</i></a></td>
							<td class="text-orange" align="right" width="180"><font size="2" face="Raleway">P&aacute;g.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
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
					consulta_exp="SELECT  'IDCliente','RazonSocial','RUC','Direccion','Telefono','Email','Activo'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="SELECT IDCliente,RazonSocial,RUC,Direccion,Telefono,Email, " & _ 
								"CASE WHEN Activo = 1 THEN 'Activo' ELSE 'Inactivo' END AS Activo " & _
								"FROM DataCRMDirconTest.dbo.Cliente " & filtrobuscador & " " & _
								"ORDER BY IDCliente" 
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt'"
					conn.execute sql

					''Tercero borrar UserExport*.xls
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"SET @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & "''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
										
					''Cuarto Uno los 2 archivos en temp*.txt
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"SET @sql='master.dbo.xp_cmdshell ''copy " & chr(34) & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt" & chr(34) & " + " & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & RutaFisicaExportar & "\UserExport" & session("codusuario") & ".xls" & chr(34) & " /b''' " & chr(10) & _
						"EXEC (@sql)"	
					conn.execute sql					
					
					''Quinto Elimino los 2 archivos en temp*.txt
					sql="DECLARE @sql VARCHAR(8000) " & chr(10) & _
						"SET @sql='master.dbo.xp_cmdshell ''del " & chr(34) & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt" & chr(34) & "," & chr(34) & RutaFisicaExportar & "\temp2_" & session("codusuario") & ".txt" & chr(34) & " " & chr(34) & "''' " & chr(10) & _
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
		alert("Ud. No tiene autorizaci�n para este proceso.");
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



