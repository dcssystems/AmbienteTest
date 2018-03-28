<%@ LANGUAGE = VBScript.Encode %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<!--El DocType es para que funcione el menu-->

<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
Dim DUSUARIO,DPASSWORD,DTELFLOGIN,DTELFPASSWORD,DCODPAIS
Dim nameUser,firsthLetter, codPerfil
	DUSUARIO		= session("goUsuario")
	DPASSWORD		= session("goClaveUsuario")
	DTELFLOGIN		= session("anexo")
	DTELFPASSWORD	= session("claveAnexo")
	DCODPAIS		= session("codigoPais")
	DTELEFONO		= session("telefono")
	nameUser        = session("nombreusuario") 
	firsthLetter    = Mid(nameUser,1,1)

if session("codusuario")<>"" then
	conectar
	sql="select C.CodPerfil,C.Descripcion,C.Orden from Usuario A inner join UsuarioPerfil B on A.codusuario=B.codusuario inner join Perfil C on B.codperfil=C.codperfil where A.codusuario=" & session("codusuario") & " and B.activo=1 UNION select A.CodPerfil,A.Descripcion,A.Orden from Perfil A where (select administrador from Usuario where codusuario=" & session("codusuario") & ")=1 order by Orden"
	consultar sql,RS1	
	if RS1.RecordCount>0 then
		codPerfil = RS1.Fields("CodPerfil")
	%>
	
	<html>
		<title><%=TITLE%></title>
		<head>
			<meta http-equiv="Content-Type" content="text/html;charset=ISO-8859-1"> 
			<link rel="stylesheet" href="assets/css/css/animation.css">
			
			<link href="https://fonts.googleapis.com/css?family=Fira+Sans+Condensed:400|Raleway:500&amp;subset=greek,greek-ext,latin-ext" rel="stylesheet"> 
			
			<link rel="shortcut icon" href="http://192.168.1.7/ambientetest/imagenes/favicon.ico">
			
			<link rel="stylesheet" href="assets/css/custom-template.css" />
			<link rel="stylesheet" href="assets/bootstrap/dist/css/bootstrap.css" />
			<script src="assets/jquery/dist/jquery.js"></script>
			<script src="assets/bootstrap/dist/js/bootstrap.js"></script>
			<script src="https://unpkg.com/sweetalert/dist/sweetalert.min.js"></script>
			
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

			
			<!--referencias menu-->
			<link rel="stylesheet" type="text/css" href="scripts/ddlevelsmenu-base.css" />
			<link rel="stylesheet" type="text/css" href="scripts/ddlevelsmenu-topbar.css" />
			<link rel="stylesheet" type="text/css" href="scripts/ddlevelsmenu-sidebar.css" />
			<script type="text/javascript" src="scripts/ddlevelsmenu.js"></script>
			<script type="text/javascript" src="scripts/popcalendar_cobcm.js"></script>
			<!--fin referencias menu-->
			
			<!--script para carpetas-->
			<script type="text/javascript" src="scripts/jquery-1.8.2.min.js"></script>
			<!--<script type="text/javascript" src="scripts/tablas.js"></script>-->
			<script type="text/javascript" language="javascript">
			window.name = "SISTEMA - CRM";
			window.status="Desarrollado por: Direct Contact Solutions"
			function agregarcol(ac_codigo,ac_descripcion,ac_url)
			{
				if ( $("#carpeta" + ac_codigo).length==0)
				{
					//alert("inicio");
					//if($('#tblTabla thead tr th').length<5)
					//{				
						//tomamos la tabla con la que vamos a trabajar
						var $objTabla=$('#tblTabla'),
						//contamos la cantidad de columnas que tiene la tabla
						iTotalColumnasExistentes=$('#tblTabla thead tr th').length;
						
						//aumentamos en uno el valor que contiene la variable
						iTotalColumnasExistentes++;
						//display:table-cell; vertical-align:middle; 
						//agregamos una columna con el titulo (en thead)
						$('<th>').html(
							'<div id="divcarpeta' + ac_codigo + '" onmousedown=func_vercarpeta("' + ac_codigo + '"); style="text-align: left; width: 178px; height: 26px; background-image: url(imagenes/carpetaon.jpg);"><img src="imagenes/vacio.png" height="7" width="178"><font face="Raleway" size="2" color="#fff">&nbsp;&nbsp;</font><a style="text-decoration: none;" href=javascript:func_vercarpeta("' + ac_codigo + '","' + ac_url + '");><font face="Raleway" size="2" color="#fff">' + ac_descripcion + '</font></a> <div style="float: right; margin-right: 5px;">'
							+
							'<a href=javascript:func_vercarpeta("' + ac_codigo + '");func_refrescarcarpeta("' + ac_codigo + '","' + ac_url + '");><i class="demo-icon2 icon-cw">&#xe80e;</i></a>'
							+
							'<a href="" id="carpeta' + ac_codigo + '" class="clsEliminar" alt="Cerrar vista" title="Cerrar vista" onclick=javascript:func_cerrarcarpeta("' + ac_codigo + '","' + ac_url + '")><i   class="demo-icon2 icon-cancel">&#xe806;</i></a></div></div>' 
							
						).appendTo($objTabla.find('thead tr'));					

						//alert("vercarpeta");
						func_vercarpeta(ac_codigo);
						//alert("refrescar url");
						//solo para cuando sea la primera ventana debemos hacer el navigate para que no loopee la barra de estado en iexplorer
						//if(iTotalColumnasExistentes==1)
						//if(navigator.appName.indexOf("Internet Explorer")!=-1) window["fr_carpeta" + ac_codigo].navigate();
						func_refrescarcarpeta(ac_codigo,ac_url);
						//alert("termine");
						//sin link '<div id="divcarpeta' + ac_codigo + '" onmousedown=func_vercarpeta("' + ac_codigo + '"); style="text-align: left; width: 178px; height: 26px; background-image: url(imagenes/carpetaon.jpg);"><img src="imagenes/vacio.png" height=4 width=178><font face="Raleway" size="2" color="#d15027">&nbsp;</font><font face="Raleway" size="2" color="#000064">' + ac_descripcion + '</font><a href=javascript:func_vercarpeta("' + ac_codigo + '");func_refrescarcarpeta("' + ac_codigo + '","' + ac_url + '");><img src="imagenes/refrescarcarpeta.png" alt="Actualizar vista" title="Actualizar vista" border=0></a><a href=""><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar" alt="Cerrar vista" title="Cerrar vista" onclick=javascript:func_cerrarcarpeta("' + ac_codigo + '","' + ac_url + '");></a></div>'
						//con link '<div id="divcarpeta' + ac_codigo + '" onmousedown=func_vercarpeta("' + ac_codigo + '"); style="text-align: left; width: 178px; height: 26px; background-image: url(imagenes/carpetaon.jpg);"><img src="imagenes/vacio.png" height=4 width=178><font face="Raleway" size="2" color="#d15027">&nbsp;</font><a style="text-decoration: none;" href=javascript:func_vercarpeta("' + ac_codigo + '","' + ac_url + '");><font face="Raleway" size="2" color="#000064">' + ac_descripcion + '</font></a><a href=javascript:func_vercarpeta("' + ac_codigo + '");func_refrescarcarpeta("' + ac_codigo + '","' + ac_url + '");><img src="imagenes/refrescarcarpeta.png" alt="Actualizar vista" title="Actualizar vista" border=0></a><a href=""><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar" alt="Cerrar vista" title="Cerrar vista" onclick=javascript:func_cerrarcarpeta("' + ac_codigo + '","' + ac_url + '");></a></div>'
						
						//'<font face="Raleway" size="2" color="#d15027">' + ac_descripcion + '</font><img id="carpeta' + ac_codigo + '" class="clsEliminar" src="imagenes/cerrarcarpeta.png" border=0>'
						//'<table cellpadding=0 cellspacing=0 border=0 height=26 width=178><THEAD><TR background="imagenes/carpetaon.jpg"><td width=178><font face"Raleway size=2 color=#d15027>'+ ac_descripcion + '</font><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar"></td></TR></THEAD></table>'
						
						//adjuntamos los td's de la columna al body de la tabla
						//$('<td>').html(
						//	'<input type="text" size="4">'
						//).appendTo($objTabla.find('tbody tr'));
						
						//cambiamos el atributo colspan del pie de la tabla y su contenido
						//$objTabla.find('tfoot tr td').attr('colspan',iTotalColumnasExistentes).
						//text('La tabla tiene '+iTotalColumnasExistentes+' columnas');
					//}
					//else alert("Cierre alguna vista para poder continuar.");
				}
				else
				{
						func_ocultartotalcarpetas();
						func_vercarpeta(ac_codigo);
				}
			}
			
			
			//clic en el enlace para eliminar la columna
				$('.clsEliminar').live('click',function(e){
					//prevenimos el comportamiento predeterminado del enlace
					e.preventDefault();
						
					//id de la tabla con la que estamos trabajando
					var $objTabla=$('#tblTabla'),
					//obtenemos el indice de la columna que se va a eliminar (padre del link)
					iColumnaAEliminar=$(this).parents('th').prevAll().length,
					//guardamos en una variable la cantidad de filas que tiene la tabla
					iTotalColumnasExistentes=$('#tblTabla thead tr th').length;
						
					//recorremos las filas de la tabla y eliminamos los td y th que se encuenten
					//en la columna que deseamos eliminar
					$objTabla.find('tr').each(function(){
						//con 'eq' especificamos el indice o la posicion del elemento
						//es como decir: eliminar el elemento TD/TH que este en el indice 4 (por ejemplo)
						$(this).find('td:eq('+iColumnaAEliminar+'),th:eq('+iColumnaAEliminar+')').remove();
					});
						
					//disminuimos la cantidad de columnas que contiene la variable (le restamos 1)
					iTotalColumnasExistentes--;
					
					/*
					//Si la eliminada es visible entonces mostramos la primera que esté disponible en imagen $("#carpeta" + ac_codigo).length==0
					//ojo que el resto ya está oculto
						var objetos = document.getElementsByTagName("iframe");
						for(var i=0; i<objetos.length; i++) {
						  var objeto = objetos[i];
						  if(objeto.name!="") 
						  {
							alert("#carpeta" + objeto.name.replace("fr_carpeta",""));
							alert($("#carpeta" + objeto.name.replace("fr_carpeta","")).length);
							if ($("#carpeta" + objeto.name.replace("fr_carpeta","")).length>0) 
							{
								document.getElementById(objeto.name).style.visibility="visible";break;
							}
						  }
						}										
					*/
					//preparamos el mensaje que vamos a mostrar en el pie de la tabla
					//var strMensaje='La tabla tiene '+iTotalColumnasExistentes+
					//((iTotalColumnasExistentes==1)?' columna':' columnas');
					//ajustamos el atributo colspan del pie de la tabla
					//$objTabla.find('tfoot tr td').attr('colspan',iTotalColumnasExistentes).text(strMensaje);
				});		
				
				function func_vercarpeta(ac_codigo)
				{
						func_ocultartotalcarpetas();
						document.getElementById("fr_carpeta" + ac_codigo).style.visibility="visible";
						document.getElementById("fr_carpeta" + ac_codigo).style.width="100%";
						document.getElementById("fr_carpeta" + ac_codigo).style.height="75%";
						//document.getElementById("fr_carpeta" + ac_codigo).style.borderColor="#8B919F";
						//document.getElementById("fr_carpeta" + ac_codigo).style.border="1";
						document.getElementById("divcarpeta" + ac_codigo).style.backgroundImage="url(imagenes/carpetaon.jpg)";
						func_apagafondoresto(ac_codigo);
				}	
				function func_refrescarcarpeta(ac_codigo,ac_url)
				{
					window.open(ac_url,"fr_carpeta" + ac_codigo);
					//window["fr_carpeta" + ac_codigo].location=ac_url;
				}
	            

				function func_cerrarcarpeta(ac_codigo)
				{
					//func_vercarpeta(ac_codigo,ac_url);
					//alert(document.getElementById("fr_carpeta" + ac_codigo).style.visibility);
					if(document.getElementById("fr_carpeta" + ac_codigo).style.visibility=="visible")
					{
						//Si la eliminada es visible entonces mostramos la primera que esté disponible en imagen $("#carpeta" + ac_codigo).length==0
						//ojo que el resto ya está oculto
						var objetos = document.getElementsByTagName("iframe");
						for(var i=0; i<objetos.length; i++) {
						  var objeto = objetos[i];
						  if(objeto.name!="") 
						  {
							//alert("#carpeta" + objeto.name.replace("fr_carpeta",""));
							//alert($("#carpeta" + objeto.name.replace("fr_carpeta","")).length);
							if ($("#carpeta" + objeto.name.replace("fr_carpeta","")).length>0 && objeto.name!="fr_carpeta" + ac_codigo) 
							{
								//document.getElementById(objeto.name).style.visibility="visible";break;
								func_vercarpeta(objeto.name.replace("fr_carpeta",""));break;
							}
						  }
						}					
					}
					document.getElementById("fr_carpeta" + ac_codigo).style.visibility="hidden";
					window.open("progvacio.html","fr_carpeta" + ac_codigo);
				}
				function func_ocultartotalcarpetas()
				{
						var objetos = document.getElementsByTagName("iframe");
						for(var i=0; i<objetos.length; i++) {
						  var objeto = objetos[i];
						  if(objeto.name!="") document.getElementById(objeto.name).style.visibility="hidden";
						}				
				}	
				function func_apagafondoresto(ac_codigo)
				{
						//Si la eliminada es visible entonces mostramos la primera que esté disponible en imagen $("#carpeta" + ac_codigo).length==0
						//ojo que el resto ya está oculto
						var objetos = document.getElementsByTagName("iframe");
						for(var i=0; i<objetos.length; i++) {
						  var objeto = objetos[i];
						  if(objeto.name!="") 
						  {
							//alert("#carpeta" + objeto.name.replace("fr_carpeta",""));
							//alert($("#carpeta" + objeto.name.replace("fr_carpeta","")).length);
							if ($("#carpeta" + objeto.name.replace("fr_carpeta","")).length>0 && objeto.name!="fr_carpeta" + ac_codigo) 
							{
								document.getElementById("divcarpeta" + objeto.name.replace("fr_carpeta","")).style.backgroundImage="url(imagenes/carpetaoff.jpg)";
								 
							}						
						  }
						}					
				}															
			</script>
			<!--fin script para carpetas-->
			
			<script type="text/javascript">
				var ventanaclave;
				function modificarclave() {
					ventanaclave = global_popup_IWTSystem(ventanaclave, "dcs_modificarclave.asp", "NewClave", "scrollbars=yes,scrolling=yes,top=" + (screen.height / 4 - 30) + ",height=160,width=450,left=" + ((screen.width - 400) / 2) + ",resizable=yes");
				}
			</script>
		</head>
		<body topmargin="0" leftmargin="0" onload="cargarLogin();">
			<table class="table-header-template" cellpadding="2" >
				<tr>
					<td><img src="imagenes/dcs_logo_agua.png" alt="Direct Contact Solutions" title="Direct Contact Solutions" height="80"></td>
					<td style="padding-right: 590px; padding-left: 150px;" class="link-text-dark" valign="middle"><font face="Raleway" size="4"><b> NOMBRE POR DEFINIR - CRM DCS </b></font></td>
					<td style="padding-left: 30px; float: right !important;" class="link-text-dark" align="right" valign="bottom">
						<table border="0" cellpadding="2" cellspacing="0">
							<tr>								
								<td>
									<div class="block">
										<div class="circle">
											<div class="dropdown dropdown-user">
												<a class="text-user dropdown-toggle" href="#" data-toggle="dropdown"><font face="Raleway" size="5" color="#fff"><%=firsthLetter%></b></font></a>												
												<ul class="dropdown-menu">
													<li><a class="link-text-dark" href="javascript:void(0);"><font face="Raleway" size="2" ><i class="demo-icon icon-up-dir">&#xe811;</i><b><%=nameUser%></b></font></a></li>
													<li><a class="link-text-dark" href="javascript:modificarclave();"><font face="Raleway" size="2" ><i class="demo-icon icon-coffee">&#xf0f4;</i> Modificar&nbsp;contrase&ntildea&nbsp;</font></a></li>
													<li class="divider"></li>
													<li><a class="link-text-dark" href="dcs_userexpira.asp" target="_top" onclick="salirLlamador();"><i class="logout demo-icon icon-logout">&#xe800;</i> Salir</a></li>
												</ul>
											</div>
										</div>
									</div>									
								</td>
							</tr>							
						</table>
					</td>
					<!--td style="padding-right:10px" id="minutoaminuto" onclick="divMarcador()"><i class="demo-icon icon-menu">&#xf008;</i></td> 0xf008 -->
					
				</tr>
			</table>
			<form name="formula" method="post">
			<table width="100%" height="24" border="0" cellpadding="2" cellspacing="0" align="center">
				<tr>
					<td style="background-color:#b72b2c">
						<!--menu inicio-->
						<div id="ddsidemenubar" class="markermenu">
							<ul>
								<li><a rel="ddsubmenuside1">MENU</a></li>
							</ul>
						</div>
						 
						<script type="text/javascript"> 
						ddlevelsmenu.setup("ddsidemenubar", "sidebar") //ddlevelsmenu.setup("mainmenuid", "topbar|sidebar")
						</script>
						<!--fin menu inicio-->
					</td>
					<td bgcolor="#b72b2c"></td>
					<td bgcolor="#b72b2c" align="right" valign="middle">
					<!--td bgcolor="#b72b2c" align="right">
						<font size="2" color="#fff"><b>Perfil:</b>&nbsp;
							<select name="codperfil" style="font-size: small; width: 170px;" onchange="document.formula.submit();">								
							<%
								codperfilselected=0
								Do while not RS1.EOF 
									if obtener("codperfil")="" and codperfilselected=0 then
										codperfilselected=RS1.Fields("CodPerfil")
									else
										if obtener("codperfil")<>"" then
											if int(obtener("codperfil"))=RS1.Fields("CodPerfil") then
												codperfilselected=RS1.Fields("CodPerfil")
											end if
										end if
									end if
								%>
									<option value=<%=RS1.Fields("CodPerfil")%> <%if codperfilselected=RS1.Fields("CodPerfil") then%> selected<%end if%>><%=RS1.Fields("Descripcion")%></option>
								<%
								RS1.MoveNext
								Loop
								RS1.Close
							%>
							</select>&nbsp;
						</font>
					</td-->
					<td align="right" bgcolor="#b72b2c"><font size="2" color="#fff"><%fechalarga=FormatDateTime(Date(),1)%><%=UCase(mid(fechalarga,1,1))%><%=mid(fechalarga,2,len(fechalarga)-1)%>&nbsp;</font></td>
				</tr>				
			</table>
			</form>
			<% if (codPerfil = 3) then %>			
			<div style="height:0px;">
				<div style="visibility:hidden;"><!-- style="visibility:hidden;" -->		
					<form name="formulaMando" id="formulaMando">
						<input name="login" id="login" type="button" value="Login" onclick="enviardatosp5('LOGIN');">
					
						<input name="logout" id="logout" type="button" value="Logout" onclick="enviardatosp5('LOGOUT')">
						
						<input name="llamar" id="llamar" type="button" value="Llamar" onclick="enviardatosp5('LLAMAR');" disabled>
						
						<center>
							<iframe src="http://192.168.1.5/p5.php" id="mandophp" name="mandophp" height="90%" width="90%"  style="height: 800px;" ></iframe>
						</center>
					</form>	
					<script language="javascript">
						var click = 1;
						async function cargarLogin()
						{		
							await sleep(2500);
							document.getElementById("login").click();
						}
						
						function salirLlamador(){
							document.getElementById("logout").click();
						}
					
						async function enviardatosp5(cadena)			
						{
							//var frame = document.getElementById('interloc'); 
							//window.parent.frames[1].contentWindow.postMessage("jajajaja", "http://192.168.1.7");
							var frameV = window.parent.frames[0].frames[1]; 
							//postMessage(mensaje, ip a donde mando el mensaje);				
							//
							switch(cadena){
								case "LOGIN":
									frameV.postMessage("E-LOGIN", "http://192.168.1.5");
									await sleep(450);
									frameV.postMessage("D-USUARIO|<%=DUSUARIO%>", "http://192.168.1.5");
									//await sleep(1);
									frameV.postMessage("D-PASSWORD|<%=DPASSWORD%>", "http://192.168.1.5");
									//await sleep(1);
									frameV.postMessage("D-TELFLOGIN|<%=DTELFLOGIN%>", "http://192.168.1.5");
									//await sleep(1);
									frameV.postMessage("D-TELFPASSWORD|<%=DTELFPASSWORD%>", "http://192.168.1.5");
									await sleep(450);
									frameV.postMessage("D-CAMP|18021801", "http://192.168.1.5"); //18021801	
									await sleep(450);
									frameV.postMessage("F-LOGIN", "http://192.168.1.5");							
								break;
								
								case "LLAMAR":
									frameV.postMessage("E-LLAMAR", "http://192.168.1.5");
									//await sleep(1);
									frameV.postMessage("D-CODPAIS|<%=DCODPAIS%>", "http://192.168.1.5");
									//await sleep(1);
									frameV.postMessage("D-TELEFONO|<%=DTELEFONO%>", "http://192.168.1.5");						
									await sleep(250);
									frameV.postMessage("F-LLAMAR", "http://192.168.1.5");							
								break;
								
								case "COLGAR":
									frameV.postMessage("E-COLGAR", "http://192.168.1.5");
								break;

								case "LOGOUT":
									frameV.postMessage("E-LOGOUT", "http://192.168.1.5");
								break;
								
							}				 
						}			
						
						function sleep(ms)
						{
						  return new Promise(resolve => setTimeout(resolve, ms));
						}
						
						function respuesta_peticion(cadena)			
						{
							switch(cadena)
							{
								case "D-LOGINOK":
									document.formulaMando.login.disabled = true;
									document.formulaMando.login.value = "LOGIN OK";
									document.formulaMando.llamar.disabled = false;
							}			
							 
						}
						
						function divMarcador()
						{
							if(click==1)
							{
							    document.getElementById("caja").style.height = "100px";
							    click = click + 1;
						    }
							else
							{
							    document.getElementById("caja").style.height = "0px";      
								click = 1;
							}  
						}
					</script>			
				</div>
			</div>
			<div id="caja">
			</div>
			
			
			<% end if %>
			<!--carpetas-->
			<TABLE id="tblTabla" class="clsTabla" cellspacing="0" cellpadding="0" border="0"><THEAD><TR></TR></THEAD></TABLE>
			<!--fin carpetas-->

			<%
			cadenaiframes=""
			cadenamenu=""
			''Armo cadenas según perfil
				''codperfilselected es el perfil seleccionado
				''si es codperfil es 1 es Perfil:Administrador entonces se mostrará todas las facultades sin excepción
				if codperfilselected=1 then
					sql="select B.CodGrupoFacultad,A.CodFacultad,B.Descripcion as GrupoFacultad,A.Descripcion as Facultad,A.pagina,B.Orden as Orden1,A.Orden as Orden2 from Facultad A inner join GrupoFacultad B on A.codgrupofacultad=B.codgrupofacultad where (select administrador from Usuario where codusuario=" & session("codusuario") & ")=1 order by B.Orden,A.Orden"
				else
					sql="select D.CodGrupoFacultad,C.CodFacultad,D.Descripcion as GrupoFacultad,C.Descripcion as Facultad,C.pagina,D.Orden as Orden1,C.Orden as Orden2 from Perfil A inner join PerfilFacultad B on A.codperfil=B.codperfil inner join Facultad C on B.codfacultad=C.CodFacultad inner join GrupoFacultad D on C.codgrupofacultad=D.codgrupofacultad where A.codperfil=" & codperfilselected & " order by D.Orden,C.Orden"
				end if
				consultar sql,RS		
				GrupoFacultad=""
				Do while not RS.EOF 
					if GrupoFacultad="" then
						GrupoFacultad=RS.Fields("GrupoFacultad")
						cadenamenu=cadenamenu & "<li><a>" & RS.Fields("GrupoFacultad") & "</a><ul>" & chr(10)
					end if
					
					if GrupoFacultad=RS.Fields("GrupoFacultad") then
						cadenamenu=cadenamenu & "<li><a href=" & chr(34) & "javascript:agregarcol('" & RS.Fields("CodFacultad") & "','" & iif(len(RS.Fields("Facultad"))<=16,RS.Fields("Facultad"),mid(RS.Fields("Facultad"),1,16) & "...") & "','" & RS.Fields("Pagina") & "');" & chr(34) & ">" & RS.Fields("Facultad") & "</a></li>" & chr(10)
					else
						GrupoFacultad=RS.Fields("GrupoFacultad")
						cadenamenu=cadenamenu & "</ul></li>" & chr(10)
						cadenamenu=cadenamenu & "<li><a>" & RS.Fields("GrupoFacultad") & "</a><ul>" & chr(10)
						cadenamenu=cadenamenu & "<li><a href=" & chr(34) & "javascript:agregarcol('" & RS.Fields("CodFacultad") & "','" & iif(len(RS.Fields("Facultad"))<=16,RS.Fields("Facultad"),mid(RS.Fields("Facultad"),1,16) & "...") & "','" & RS.Fields("Pagina") & "');" & chr(34) & ">" & RS.Fields("Facultad") & "</a></li>" & chr(10)
					end if
					
					cadenaiframes=cadenaiframes & "<IFRAME SRC='progvacio.html' style='visibility: hidden; position: absolute; overflow-x: hidden;' allowTransparency='true' HSPACE=0 ALIGN=TOP WIDTH=0 HEIGHT=0 FRAMEBORDER=0 NAME='fr_carpeta" & RS.Fields("CodFacultad") & "' ID='fr_carpeta" & RS.Fields("CodFacultad") & "' SCROLLING=AUTO><BLOCKQUOTE><P>Debe utilizar IExplorer 5.5 o superior.</P></BLOCKQUOTE></IFRAME>" & chr(10)
					
				RS.MoveNext
				Loop
				RS.Close				
				cadenamenu=cadenamenu & "</ul></li>" & chr(10)
			%>

			<!--visibilidad de carpetas-->
			<!--iframes vacios por la cantidad de facultades del perfil: Perfil Facultad-->
			<%=cadenaiframes%>
			<!--fin visibildad carpetas-->

			<!--datos del menu:  Inicio-->
			<ul id="ddsubmenuside1" class="imagenes/arrow-right.gif blackwhite">
			<%=cadenamenu%>
			<!--fin datos del menu:  Inicio-->		

			
		</body>
		
	</html>
	<%
	else
	%>
	<script language="javascript">
		alert("Ud. No tiene autorización para este proceso.");
		window.open("dcs_userexpira.asp","_top");
	</script>	
	<%
	end if
	desconectar
else
%>
<script language="javascript">
	alert("Finalizo la sesión en el sistema. Ingrese al sistema");
	window.open("index.html","_top");
</script>
<%
end if
%>



