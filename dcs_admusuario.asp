<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
''NOTA: Si se le da a un usuario el Perfil:Administrador por defecto se activa el flag: administrador y viceversa
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admusuario.asp") then
		buscador=obtener("buscador")
		buscaractivos=obtener("buscaractivos")
		buscarbloqueados=obtener("buscarbloqueados")
		buscaradministrador=obtener("buscaradministrador")
		
		
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
		<!--cargando--><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
		<head>
		<script language=javascript src="scripts/TablaDinamica.js"></script>
		<script type="text/javascript" src="scripts/tristate-0.9.2.js" ></script>
		<script language=javascript>
		var ventanauser;
		function inicio()
		{
		dibujarTabla(0);
		}
		function modificar(codigo)
		{
			ventanauser=global_popup_IWTSystem(ventanauser,"nuevousuario.asp?vistapadre=" + window.name + "&paginapadre=admusuario.asp&codusuario=" + codigo,"NewUser","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 425)/2 - 30) + ",height=445,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
		}			
		function agregar()
		{
			ventanauser=global_popup_IWTSystem(ventanauser,"nuevousuario.asp?vistapadre=" + window.name + "&paginapadre=admusuario.asp","NewUser","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 425)/2 - 30) + ",height=445,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
		  //Configuraci�n general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=2;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=14;
		    fondovariable[tabla]='bgcolor=#f5f5f5';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('codusuario','Usuario','Ap.Paterno','Ap.Materno','Nombres','e-Mail','Territorio','Oficina','Agencia','Tipo','Act','Bloq','Adm','Editar');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(false, true, true ,true , true , true , true, true  , true , true , true  , true  , true   , true );
		    anchocolumna[tabla] =  new Array( '',     '3%', '3%' , '3%' ,    '3%',    '4%' ,   '5%' ,    '5%',    '3%', '3%' , '3%'  , '3%' ,'3%','2%');
		    aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left','left','left' ,'center'  , 'center','center','left');
		    aligndetalle[tabla] = new Array('left','left','left','left','left','left','left','left','left','left' ,'center'  , 'center','center','left');
		    alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left','left' ,'left'  , 'left','left','left');
		    decimalesnumero[tabla] = new Array(-1 ,-1   ,-1      ,-1   ,-1 ,-1  ,-1    ,-1    ,-1    ,-1 ,-1    ,-1      ,-1,-1);
		    formatofecha[tabla] =   new Array(''  ,''   ,''      ,''  ,'' ,'','','',   '', '', '',  '',     '',     '');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
				objetofomulario[tabla][0]='';
				objetofomulario[tabla][1]='<input type=hidden name=codusuario-id- value=-c0->' + '<a href="javascript:modificar(-id-);">-valor-</a>';//+ objetodatos("text",tabla,"usuario","left","6","");
				objetofomulario[tabla][2]=objetodatos("text",tabla,"apepaterno","left","20","");
				objetofomulario[tabla][3]=objetodatos("text",tabla,"apematerno","left","20","");
				objetofomulario[tabla][4]=objetodatos("text",tabla,"nombres","left","20","");
				objetofomulario[tabla][5]='<a href="javascript:modificar(-id-);">-valor-</a>';//objetodatos("text",tabla,"correo","left","22","");
				objetofomulario[tabla][6]='<a href="javascript:modificar(-id-);">-valor-</a>';//territorio			
				objetofomulario[tabla][7]='<a href="javascript:modificar(-id-);">-valor-</a>';//oficina 
				//Todo select me devuel: '<select name="' + nombre + '-id-" style="font-size: xx-small;" onchange="datos[-t-][-i-][-j-]=this.options[this.selectedIndex].text;"></select>';
				//onchange="this.title=this.options[this.selectedIndex].text;this.alt=this.options[this.selectedIndex].text;"
				//onmouseover="this.title=' + String.fromCharCode(39) + 'Agencia Seleccionada: ' + String.fromCharCode(39) + ' + this.options[this.selectedIndex].text;this.alt=' + String.fromCharCode(39) + 'Agencia Seleccionada: ' + String.fromCharCode(39) + ' + this.options[this.selectedIndex].text;"
				//--objetofomulario[tabla][8]=objetodatos("select",tabla,"codagencia","","","").replace('style="font-size: xx-small;"','style="font-size: xx-small;width: 150px;"');
				objetofomulario[tabla][8]='<a href="javascript:modificar(-id-);">-valor-</a>';//agencia
				objetofomulario[tabla][9]='<a href="javascript:modificar(-id-);">-valor-</a>';//codtipousuario
				objetofomulario[tabla][10]=objetodatos("checkbox",tabla,"activo","","","");
				objetofomulario[tabla][11]=objetodatos("checkbox",tabla,"fbloq","","","");
				objetofomulario[tabla][12]=objetodatos("checkbox",tabla,"administrador","","","");
				objetofomulario[tabla][13]='<a href="javascript:modificar(-id-);">Editar</a>';
										
					
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
				filtrofomulario[tabla][8]='';
				filtrofomulario[tabla][9]='';
				filtrofomulario[tabla][10]='';
				filtrofomulario[tabla][11]='';
				filtrofomulario[tabla][12]='';
				filtrofomulario[tabla][13]='';
										
					
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
				


		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		if buscador<>"" then
			filtrobuscador = " where (A.usuario like '%" & buscador & "%' or A.apepaterno like '%" & buscador & "%' or A.apematerno like '%" & buscador & "%' or A.nombres like '%" & buscador & "%' or A.correo like '%" & buscador & "%' or B.razonsocial like '%" & buscador & "%' or C.codoficina + ' - ' +  C.descripcion like '%" & buscador & "%' or D.codterritorio + ' - ' + D.descripcion like '%" & buscador & "%' or E.Descripcion like '%" & buscador & "%') "
		end if
		''if buscaractivos<>"" then
		''	filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.activo=1"
		''end if		
		
		select case buscaractivos
		case "0" : 
					checkbuscactivos="value='0'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.activo=0"
		case "2" : 
					checkbuscactivos="value='2'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.activo=1"
		case else: 
					checkbuscactivos="value='1'"
		end select		
		
		select case buscarbloqueados
		case "0" : 
					checkbuscbloqueados="value='0'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.flagbloqueo=0"
		case "2" : 
					checkbuscbloqueados="value='2'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.flagbloqueo=1"
		case else: 
					checkbuscbloqueados="value='1'"
		end select		
		
		select case buscaradministrador
		case "0" : 
					checkbuscadministrador="value='0'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.administrador=0"
		case "2" : 
					checkbuscadministrador="value='2'"
					filtrobuscador = filtrobuscador & iif(filtrobuscador=""," where "," and ") & "A.administrador=1"
		case else: 
					checkbuscadministrador="value='1'"
		end select		
		
								
				
		contadortotal=0
		sql="select count(*) from usuario A left outer join agencia B on A.codagencia=B.codagencia left outer join oficina C on A.codoficina=C.codoficina left outer join territorio D on C.codterritorio=D.codterritorio left outer join TipoUsuario E on A.codtipousuario=E.codtipousuario " & filtrobuscador 
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
		sql="select top " & cantidadxpagina & " A.*,B.razonsocial as agencia,D.codterritorio + ' - ' + D.descripcion as territorio,C.codoficina + ' - ' + C.descripcion as oficina,CASE WHEN FlagBloqueo<3 THEN 0 ELSE 1 END as FBloq,E.Descripcion as TipoUsuario " & _
			"from usuario A left outer join agencia B on A.codagencia=B.codagencia left outer join oficina C on A.codoficina=C.codoficina left outer join territorio D on C.codterritorio=D.codterritorio left outer join TipoUsuario E on A.codtipousuario=E.codtipousuario where " & filtrobuscador1 & " A.codusuario not in (select top " & topnovisible & " A.codusuario from usuario A left outer join agencia B on A.codagencia=B.codagencia left outer join oficina C on A.codoficina=C.codoficina left outer join territorio D on C.codterritorio=D.codterritorio left outer join TipoUsuario E on A.codtipousuario=E.codtipousuario " & filtrobuscador & " order by A.apepaterno,A.apematerno,A.nombres,A.codusuario) order by A.apepaterno,A.apematerno,A.nombres,A.codusuario" 
		else
		sql="select top " & cantidadxpagina & " A.*,B.razonsocial as agencia,D.codterritorio + ' - ' + D.descripcion as territorio,C.codoficina + ' - ' + C.descripcion as oficina,CASE WHEN FlagBloqueo<3 THEN 0 ELSE 1 END as FBloq,E.Descripcion as TipoUsuario " & _
			"from usuario A left outer join agencia B on A.codagencia=B.codagencia left outer join oficina C on A.codoficina=C.codoficina left outer join territorio D on C.codterritorio=D.codterritorio left outer join TipoUsuario E on A.codtipousuario=E.codtipousuario " & filtrobuscador & " order by A.apepaterno,A.apematerno,A.nombres,A.codusuario" 
		end if
		''response.write sql
		consultar sql,RS
		contador=0
		
			Do while not RS.EOF
				if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then
					if obtener("apepaterno" & RS.Fields("codusuario"))<>"" then
						apepaterno=obtener("apepaterno" & RS.Fields("codusuario"))
					else
						apepaterno=""
					end if
					if obtener("apematerno" & RS.Fields("codusuario"))<>"" then
						apematerno=obtener("apematerno" & RS.Fields("codusuario"))
					else
						apematerno=""
					end if	
					if obtener("nombres" & RS.Fields("codusuario"))<>"" then
						nombres=obtener("nombres" & RS.Fields("codusuario"))
					else
						nombres=""
					end if	
																									
					''''codagencia=obtener("codagencia" & RS.Fields("codusuario"))
					''''if not isNumeric(codagencia) then
					''''	codagencia="0"
					''''end if		
					
					if obtener("fbloq" & RS.Fields("codusuario"))<>"" then
						fbloq="1"
					else
						fbloq="0"
					end if
					if obtener("activo" & RS.Fields("codusuario"))<>"" then
						activo="1"
					else
						activo="0"
					end if						
					if obtener("administrador" & RS.Fields("codusuario"))<>"" then
						administrador="1"
					else
						administrador="0"
					end if	
										
						''obtener("correo" & RS.Fields("codusuario")) <> rs.Fields("correo") or _
						''int(codagencia) <> iif(IsNull(rs.Fields("codagencia")),0,rs.Fields("codagencia")) or _
						
						if 	obtener("apepaterno" & RS.Fields("codusuario")) <> rs.Fields("apepaterno") or _
							obtener("apematerno" & RS.Fields("codusuario")) <> rs.Fields("apematerno") or _
							obtener("nombres" & RS.Fields("codusuario")) <> rs.Fields("nombres") or _
							int(activo)<> rs.Fields("activo") or _
							int(fbloq) <> rs.fields("fbloq") or _
							int(administrador) <> rs.Fields("administrador") then
								if fbloq = "1" then
								xfbloq = "3"
								else
								xfbloq = "0"
								end if
								
							existeotroadmin=0
							if rs.Fields("administrador")=0 and administrador="1" then
								''antes de insertarlo activo si existiera
								sql="Update UsuarioPerfil set activo=1,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codperfil=1 and codusuario=" & rs.Fields("codusuario")
								conn.Execute sql						
								''lo inserta si no existe		
								sql="insert into UsuarioPerfil (codusuario,codperfil,usuarioregistra,fecharegistra,activo) select " & rs.Fields("codusuario") & ",1," & session("codusuario") & ",getdate(),1 where (select count(*) from UsuarioPerfil where codusuario=" & rs.Fields("codusuario") & " and codperfil=1)=0"
								conn.Execute sql
							end if
							if rs.Fields("administrador")=1 and administrador="0" then
								sql="select count(*) from usuario where administrador=1 and codusuario<>" & rs.Fields("codusuario")
								consultar sql,RS1
								existeotroadmin=RS1.fields(0)
								RS1.Close
								if existeotroadmin=0 then
									administrador="1"
								else
									sql="Update UsuarioPerfil set activo=0,usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codperfil=1 and codusuario=" & rs.Fields("codusuario")
									conn.Execute sql
								end if								
							end if
							
							''''if codagencia="0" then
							''''	codagenciagrab="null"
							''''else
							''''	codagenciagrab=codagencia
							''''end if

							sql="update usuario set apepaterno='" & apepaterno & "',apematerno='" & apematerno & "',nombres='" & nombres & "',activo=" & activo & ",flagbloqueo=" & xfbloq & ",administrador=" & administrador & ",usuariomodifica=" & session("codusuario") & ",fechamodifica=getdate() where codusuario=" & rs.Fields("codusuario") 
							''response.write sql
							conn.Execute sql

						end if
						
								
													
						''''sql="select Descripcion from agencia where codagencia=" & codagencia
						''''consultar sql,RS1
						''''if not RS1.eof then
						''''nomagencia=RS1.fields(0)
						''''else
						''''nomagencia = ""
						''''end if
						''''RS1.Close
					
					
					end if 

										
		%>
			datos[tabla][<%=contador%>] = new Array();
				datos[tabla][<%=contador%>][0]=<%=RS.Fields("codusuario")%>;
				//datos[tabla][<%=contador%>][1]='<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%=usuario%><%else%><%=rs.Fields("Usuario")%><%end if%>';
				datos[tabla][<%=contador%>][1]='<%=rs.Fields("Usuario")%>';
			    datos[tabla][<%=contador%>][2]='<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%=apepaterno%><%else%><%=rs.Fields("apepaterno")%><%end if%>';
				datos[tabla][<%=contador%>][3]='<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%=apematerno%><%else%><%=rs.Fields("apematerno")%><%end if%>';
				datos[tabla][<%=contador%>][4]='<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%=nombres%><%else%><%=rs.Fields("nombres")%><%end if%>';
				datos[tabla][<%=contador%>][5]='<%=rs.Fields("correo")%>';				
				datos[tabla][<%=contador%>][6]='<%=replace(iif(IsNull(rs.Fields("Territorio")),"",rs.Fields("Territorio"))," ","&nbsp;")%>'
				datos[tabla][<%=contador%>][7]='<%=replace(iif(IsNull(rs.Fields("Oficina")),"",rs.Fields("Oficina"))," ","&nbsp;")%>';
				//datos[tabla][<%=contador%>][8]='<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%=nomagencia%><%else%><%=rs.Fields("agencia")%><%end if%>';
				datos[tabla][<%=contador%>][8]='<%=replace(iif(IsNull(rs.Fields("Agencia")),"",rs.Fields("Agencia"))," ","&nbsp;")%>';
				datos[tabla][<%=contador%>][9]='<%=rs.Fields("TipoUsuario")%>';
				datos[tabla][<%=contador%>][10]=<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%if int(activo)=1 then%>'checked'<%else%>' '<%end if%><%else%><%if rs.Fields("activo")=1 then%>'checked'<%else%>' '<%end if%><%end if%>;
				datos[tabla][<%=contador%>][11]=<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%if int(fbloq)=1 then%>'checked'<%else%>' '<%end if%><%else%><%if rs.Fields("fbloq")=1 then%>'checked'<%else%>' '<%end if%><%end if%>;
				datos[tabla][<%=contador%>][12]=<%if obtener("actualizarlista")<>"" and obtener("codusuario" & RS.Fields("codusuario"))<>"" then%><%if int(administrador)=1 then%>'checked'<%else%>' '<%end if%><%else%><%if rs.Fields("administrador")=1 then%>'checked'<%else%>' '<%end if%><%end if%>;
				datos[tabla][<%=contador%>][13]='';
		<%
			contador=contador + 1
			RS.MoveNext 
			Loop 
			RS.Close
		%>
			    
		    //datos del pie si fuera visible
		    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
		    piefunciones[tabla] = new Array('','','','','','','','','','','','','',''); 


		    //Se escriben las opciones para los selects que contenga
		    posicionselect[tabla]=new Array();
		    nombreselect[tabla]=new Array();
		    opcionesvalor[tabla]=new Array();
		    opcionestexto[tabla]=new Array();
		    //Finaliza configuracion de tabla 0
		    
		    /*
			posicionselect[tabla][0]=8; //columna del select
			nombreselect[tabla][0]="codagencia"; //nombre del select
			
			var miArrayOPValor = new Array();
			var miArrayOPTexto = new Array();

			<%
			''''contadorlinea=0
			''''sql="select 0 as codagencia,'' as Descripcion UNION select codagencia,Descripcion from agencia order by Descripcion"
			''''''sql="select codagencia,Descripcion from agencia order by Descripcion"
			''''consultar sql,RS
			''''Do While not RS.EOF
				%>
				miArrayOPValor[<%''''=contadorlinea%>]=<%''''=RS.fields("codagencia")%>;
				miArrayOPTexto[<%''''=contadorlinea%>]='<%''''=RS.fields("Descripcion")%>';
				<%
			''''RS.MoveNext
			''''contadorlinea=contadorlinea + 1
			''''Loop
			''''RS.Close
			%>
			opcionesvalor[tabla][0] = miArrayOPValor;
			opcionestexto[tabla][0] = miArrayOPTexto;		
			*/
		    funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
		    funcionagrega[tabla]='agregar();';

		</script> 
		
		<%
        objetosdebusqueda="<font size=2 face=Arial color=#00529B>Buscar:&nbsp;<input name='buscador' value='" & buscador & "' size=20 onkeypress='if(window.event.keyCode==13) buscar();'></font>&nbsp;<span id='tristateBox1' style='cursor: default;'>&nbsp;Activos<input type='hidden' id='tristateBox1State' name='buscaractivos' " & checkbuscactivos & "></span>&nbsp;<span id='tristateBox2' style='cursor: default;'>&nbsp;Bloqueados<input type='hidden' id='tristateBox2State' name='buscarbloqueados' " & checkbuscbloqueados & "></span>&nbsp;<span id='tristateBox3' style='cursor: default;'>&nbsp;Administrador<input type='hidden' id='tristateBox3State' name='buscaradministrador' " & checkbuscadministrador & "></span>"
		%>	
		
		<%if contador=0 then%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0>	
			<tr>
				<td bgcolor="#F5F5F5"><font size=2 face=Arial color=#00529B><b>Usuarios (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
			</tr>
			</table>
		<%else		
		%>
			<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF"><!--onload="inicio();"-->
			<form name=formula method=post>
			<table width=100% cellpadding=4 cellspacing=0 border=0>		
			<tr>
				<td bgcolor="#F5F5F5" align=left><font size=2 face=Arial color=#00529B><b>Usuarios (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:actualizar();"><img src="imagenes/guardar.gif" border=0 alt="Guardar" title="Guardar" align=middle></a>&nbsp;&nbsp;<a href="javascript:agregar();"><img src="imagenes/nuevo.gif" border=0 alt="Nuevo" title="Nuevo" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
				<td bgcolor="#F5F5F5" align=right><%=objetosdebusqueda%></td>
				<td bgcolor="#F5F5F5" align=left><a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
				<td bgcolor="#F5F5F5" align=right width=180><font size=2 face=Arial color=#00529B>P�g.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
			</tr>	
			</table>
			<div id="tabla0"> 
			</div>
		<%end if%>
		<input type="hidden" name="actualizarlista" value="">
		<%''Codigo exp excel%>
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="pag" value="<%=pag%>">	
		</form>
		<script type="text/javascript">
			initTriStateCheckBox('tristateBox1', 'tristateBox1State', true);
		</script>
		<script type="text/javascript">
			initTriStateCheckBox('tristateBox2', 'tristateBox2State', true);
		</script>
		<script type="text/javascript">
			initTriStateCheckBox('tristateBox3', 'tristateBox3State', true);
		</script>
		<script language="javascript">
			inicio();
		</script>					
		</body>
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display="none";</script>
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
					consulta_exp="select 'Usuario','Ap.Paterno','Ap.Materno','Nombres','e-Mail','Territorio','Oficina','Agencia','Tipo','Activo','Bloqueo','Administrador'"
					sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
					conn.execute sql
					
					''Segundo Detalle en temp2_(user).txt
					consulta_exp="select A.Usuario,A.apepaterno,A.apematerno,A.nombres,A.correo,D.codterritorio + ' - ' + D.descripcion as territorio,C.codoficina + ' - ' + C.descripcion as oficina,B.razonsocial as agencia,E.Descripcion as TipoUsuario,CASE WHEN A.Activo=1 THEN 'S�' ELSE 'No' END as FlagActivo,CASE WHEN A.FlagBloqueo<3 THEN 'No' ELSE  'S�' END as FlagBloqueo,CASE WHEN A.Administrador=1 THEN 'S�' ELSE  'No' END as FlagAdmin " & _
								 "from CobranzaCM.dbo.usuario A left outer join CobranzaCM.dbo.agencia B on A.codagencia=B.codagencia left outer join CobranzaCM.dbo.oficina C on A.codoficina=C.codoficina left outer join CobranzaCM.dbo.territorio D on C.codterritorio=D.codterritorio left outer join CobranzaCM.dbo.TipoUsuario E on A.codtipousuario=E.codtipousuario " & filtrobuscador & " order by A.apepaterno,A.apematerno,A.nombres,A.codusuario" 
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
		alert("Ud. No tiene autorizaci�n para este proceso.");
		window.open("dcs_userexpira.asp","_top");
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



