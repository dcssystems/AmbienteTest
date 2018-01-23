<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admfacultad.asp") then
		buscador=obtener("buscador")
		'expimp=obtener("expimp")
	%>
		<html>
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<script language=javascript src="scripts/TablaDinamica.js"></script>
		<script language=javascript>
		var ventanauser;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanauser=global_popup_IWTSystem(ventanauser,"nuevofacultad.asp?vistapadre=" + window.name + "&paginapadre=admusuario.asp&codusuario=" + codigo,"NewUser","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 425)/2 - 30) + ",height=445,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregar()
		{
			ventanauser=global_popup_IWTSystem(ventanauser,"nuevofacultad.asp?visortabla=admusuarios","NewUser","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 425)/2 - 30) + ",height=445,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
		<style>
		A {
			FONT-SIZE: 12px; COLOR: #00529B; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
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
			font-size:12px;
			font-family:Arial;
			cursor:hand;
		}
		TD
		{
			color:#00529B;
			font-size:12px;
			font-family:Arial;
		}
		TR
		{
			background: #FFFFFF;
		}
		</style>
		</head>
		
		<script language=javascript>
			rutaimgcab="imagenes/"; 
		  //Configuración general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=6;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('Codigo','Grupo','Descripción','Pagina','Orden','');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(true, true, true ,true,true, true);
		    anchocolumna[tabla] =  new Array( '1%',     '1%', '1%' , '1%','1%' ,'');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1   ,-1      ,-1 ,-1 , -1  );
		    formatofecha[tabla] =   new Array(''  ,''   ,''      ,'' ,'','' );


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='<input type=hidden name=codfacultad-id- value=-c0->' + '<a href="javascript:modificar(-id-);">-valor-</a>';
				objetofomulario[tabla][1]=objetodatos("select",tabla,"codgrupofacultad","","","").replace('style="font-size: xx-small;"','style="font-size: xx-small;width: 150px;"');
				objetofomulario[tabla][2]=objetodatos("text",tabla,"descripcion","left","22","");
				objetofomulario[tabla][3]=objetodatos("text",tabla,"pagina","left","22","");
				objetofomulario[tabla][4]=objetodatos("text",tabla,"orden","left","5","");
				objetofomulario[tabla][5]='';
				
										
					
		    filtrardatos[tabla]=0; //define si carga auto el filtro
		    filtrofomulario[tabla] = new Array();
		    tipofiltrofomulario[tabla] = new Array();
		    	filtrofomulario[tabla][0]='';
				filtrofomulario[tabla][1]='';
				filtrofomulario[tabla][2]=''; //objetofiltro("text",tabla,2,'contiene');
				filtrofomulario[tabla][3]='';
				filtrofomulario[tabla][4]='';
				filtrofomulario[tabla][5]='';
				
				
										
					
		    valorfiltrofomulario[tabla] = new Array();
				valorfiltrofomulario[tabla][0]='';
				valorfiltrofomulario[tabla][1]='';
				valorfiltrofomulario[tabla][2]='';
				valorfiltrofomulario[tabla][3]='';
				valorfiltrofomulario[tabla][4]='';
				valorfiltrofomulario[tabla][5]='';


		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		if buscador<>"" then
			filtrobuscador = " where (f.codfacultad like '%" & buscador & "%' or f.descripcion like '%" & buscador & "%' or g.descripcion like '%" & buscador & "%' or f.orden like '%" & buscador & "%') "
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

		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if
		
		if pag>1 then					
		sql="select top " & cantidadxpagina & " *,f.descripcion as descrip,g.descripcion as grupo, f.orden from facultad f inner join grupofacultad g on f.codgrupofacultad = g.codgrupofacultad where " & filtrobuscador1 & " f.codfacultad not in (select top " & topnovisible & " f.codfacultad from facultad f inner join grupofacultad g on f.codgrupofacultad = g.codgrupofacultad " & filtrobuscador & " order by f.codfacultad) order by f.codfacultad" 
		else
		sql="select top " & cantidadxpagina & " *,f.descripcion as descrip,g.descripcion as grupo, f.orden from facultad f inner join grupofacultad g on f.codgrupofacultad = g.codgrupofacultad " & filtrobuscador & " order by f.codfacultad" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF 
				if obtener("actualizarlista")<>"" and obtener("codfacultad" & RS.Fields("codfacultad"))<>"" then
					if obtener("descripcion" & RS.Fields("codfacultad"))<>"" then
						descripcion=obtener("descripcion" & RS.Fields("codfacultad"))
					else
						descripcion=""
					end if		
					if obtener("pagina" & RS.Fields("codfacultad"))<>"" then
						pagina=obtener("pagina" & RS.Fields("codfacultad"))
					else
						pagina=""
					end if	
					if obtener("orden" & RS.Fields("codfacultad"))<>"" then
						orden=obtener("orden" & RS.Fields("codfacultad"))
					else
						orden=""
					end if											
					codgrupofacultad=obtener("codgrupofacultad" & RS.Fields("codfacultad"))
					if not isNumeric(codgrupofacultad) then
						codgrupofacultad="0"
					end if			
						if 	obtener("descripcion" & RS.Fields("codfacultad")) <> rs.Fields("descripcion") or _
							int(codgrupofacultad) <> iif(IsNull(rs.Fields("codgrupofacultad")),0,rs.Fields("codgrupofacultad")) or _
							obtener("pagina" & RS.Fields("codfacultad")) <> rs.Fields("pagina") then
															
							
							sql="update facultad set descripcion='" & descripcion & "',pagina='" & pagina & "',codgrupofacultad='" & codgrupofacultad & "',usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate(), orden ="& orden &" where codfacultad=" & rs.Fields("codfacultad") 
							''response.write sql
							conn.Execute sql

						end if
						
								
													
						sql="select g.descripcion   from grupofacultad as g where codgrupofacultad =" & codgrupofacultad
						consultar sql,RS1
						if not RS1.eof then
						nomgrupo=RS1.fields(0)
						else
						nomgrupo = ""
						end if
						RS1.Close
					
					
				end if 

										
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]=<%=RS.Fields("codfacultad")%>;
				datos[tabla][<%=contador%>][1]='<%if obtener("actualizarlista")<>"" and obtener("codfacultad" & RS.Fields("codfacultad"))<>"" then%><%=nomgrupo%><%else%><%=rs.Fields("grupo")%><%end if%>';
				datos[tabla][<%=contador%>][2]='<%if obtener("actualizarlista")<>"" and obtener("codfacultad" & RS.Fields("codfacultad"))<>"" then%><%=descripcion%><%else%><%=rs.Fields("descrip")%><%end if%>';
				datos[tabla][<%=contador%>][3]='<%if obtener("actualizarlista")<>"" and obtener("codfacultad" & RS.Fields("codfacultad"))<>"" then%><%=pagina%><%else%><%=rs.Fields("pagina")%><%end if%>';
				datos[tabla][<%=contador%>][4]='<%if obtener("actualizarlista")<>"" and obtener("codfacultad" & RS.Fields("codfacultad"))<>"" then%><%=orden%><%else%><%=rs.Fields("orden")%><%end if%>';
				datos[tabla][<%=contador%>][5]='';
							
		<%
			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','',''); 


		    //Se escriben las opciones para los selects que contenga
		    posicionselect[tabla]=new Array();
		    nombreselect[tabla]=new Array();
		    opcionesvalor[tabla]=new Array();
		    opcionestexto[tabla]=new Array();
		    //Finaliza configuracion de tabla 0
		    
			posicionselect[tabla][0]=1; //columna del select
			nombreselect[tabla][0]="codgrupofacultad"; //nombre del select
			
			var miArrayOPValor = new Array();
			var miArrayOPTexto = new Array();

			<%
			contadorlinea=0
			sql="select 0 as codgrupofacultad,'' as Descripcion UNION select codgrupofacultad,Descripcion from grupofacultad order by Descripcion"
			''sql="select codagencia,Descripcion from agencia order by Descripcion"
			consultar sql,RS
			Do While not RS.EOF
				%>
				miArrayOPValor[<%=contadorlinea%>]=<%=RS.fields("codgrupofacultad")%>;
				miArrayOPTexto[<%=contadorlinea%>]='<%=RS.fields("Descripcion")%>';
				<%
			RS.MoveNext
			contadorlinea=contadorlinea + 1
			Loop
			RS.Close
			%>
			opcionesvalor[tabla][0] = miArrayOPValor;
			opcionestexto[tabla][0] = miArrayOPTexto;		
			    
		    funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
		    funcionagrega[tabla]='agregar();';

		</script> 
		
		<%
		
				if expimp="1" then
					''Para Exportar a Excel
					excel=""
					excel=excel & "Reporte de Usuarios" & chr(10)
					excel=excel & "Usuario" & chr(9) & "Nombres" & chr(9) & "Cargo" & chr(9) & "Alm/Tda" & chr(9) & "F.Nac." & chr(9) & "Teléfono" & chr(9) & "E-Mail" & chr(9) & "Activo" & chr(9) & "Administrador" & chr(9) & "Ventas" & chr(9) & "Compras" & chr(9) & "Almacén" & chr(9) & "Cuentas" & chr(9) & "Reportes" & chr(9) & "Eliminaciones" & chr(9) & "Anulaciones" & chr(9) & "Modifica Tipo Cambio" & chr(9) & "Modifica Créditos" & chr(9) & "Modifica Precios" & chr(10)
						sql="select *,(select Descripcion from Almacen where IDAlmacen=Usuario.IDAlmacen) as NomAlmacen from Usuario order by activo desc,nombres" 
						consultar sql,RS
						Do While not RS.EOF
							if not IsNull(RS.Fields("FechaNac")) then
								excel=excel & RS.Fields("Usuario") & chr(9) & RS.Fields("Nombres") & chr(9) & RS.Fields("Cargo") & chr(9) & RS.Fields("NomAlmacen") & chr(9) & day(RS.Fields("FechaNac")) & "/" & month(RS.Fields("FechaNac")) & "/" & year(RS.Fields("FechaNac")) & chr(9) & RS.Fields("Telefono") & chr(9) & RS.Fields("EMail") & chr(9)
							else
								excel=excel & RS.Fields("Usuario") & chr(9) & RS.Fields("Nombres") & chr(9) & RS.Fields("Cargo") & chr(9) & RS.Fields("NomAlmacen") & chr(9) & chr(9) & RS.Fields("Telefono") & chr(9) & RS.Fields("EMail") & chr(9)
							end if
							if RS.Fields("Activo")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if
							if RS.Fields("Administrador")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if
							if RS.Fields("Ventas")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if				
							if RS.Fields("Compras")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if				
							if RS.Fields("Almacen")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if				
							if RS.Fields("Cuentas")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if											
							if RS.Fields("Reportes")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if																	
							if RS.Fields("Eliminaciones")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if						
							if RS.Fields("Anulaciones")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if					
							if RS.Fields("PrivTC")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if		
							if RS.Fields("ModCredito")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if	
							if RS.Fields("ModPrecio")=1 then
							excel=excel & "x" & chr(9) 
							else
							excel=excel & "" & chr(9) 
							end if																									
							excel=excel & chr(10)
						RS.MoveNext
						Loop
						RS.Close
					nombrearchivo=ExportarExcel(excel)
				%>
					<script language="javascript">
					window.open('exportados/<%=nombrearchivo%>.xls','VerExport');
					</script>
				<%					
				end if	
		%>	
		
		<%if contador=0 then%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Facultad (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=middle width=200><font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size=20 onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
			</tr>
			</table>
		<%else		
		%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Facultad (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar" align=middle></a>&nbsp;&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a>&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a><%if expimp="1" then%>&nbsp;&nbsp;<a href='exportados/<%=nombrearchivo%>.xls','VerExport'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=middle width=200><font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size=20 onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
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



