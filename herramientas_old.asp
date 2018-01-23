<%@ LANGUAGE = VBScript.Encode %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<!--El DocType es para que funcione el menu-->

<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
%>
	<html>
		<title>Sistema Web de Gestión de Cobranzas BBVA</title>
		<head>
		<!--referencias menu-->
		<link rel="stylesheet" type="text/css" href="scripts/ddlevelsmenu-base.css" />
		<link rel="stylesheet" type="text/css" href="scripts/ddlevelsmenu-topbar.css" />
		<link rel="stylesheet" type="text/css" href="scripts/ddlevelsmenu-sidebar.css" />
		<script type="text/javascript" src="scripts/ddlevelsmenu.js"></script>
		<!--fin referencias menu-->
		
		<!--script para carpetas-->
		<script type=text/javascript src="scripts/jquery-1.8.2.min.js"></script>
		<!--<script type=text/javascript src="scripts/tablas.js"></script>-->
		<script type=text/javascript>
		function agregarcol(ac_codigo,ac_descripcion,ac_url)
		{
			if ( $("#carpeta" + ac_codigo).length==0)			{				//if($('#tblTabla thead tr th').length<5)
				//{				
					//tomamos la tabla con la que vamos a trabajar
					var $objTabla=$('#tblTabla'),
					//contamos la cantidad de columnas que tiene la tabla
					iTotalColumnasExistentes=$('#tblTabla thead tr th').length;
					
					//aumentamos en uno el valor que contiene la variable
					iTotalColumnasExistentes++;
					//display:table-cell; vertical-align:middle; 
					//agregamos una columna con el titulo (en thead)					$('<th>').html(
						'<div id="divcarpeta' + ac_codigo + '" onmousedown=func_vercarpeta("' + ac_codigo + '"); style="text-align: left; width: 178px; height: 26px; background-image: url(imagenes/carpetaon.jpg);"><img src="imagenes/vacio.png" height=7 width=178><font face="Arial" size="2" color="#00529B">&nbsp;&nbsp;</font><a style="text-decoration: none;" href=javascript:func_vercarpeta("' + ac_codigo + '","' + ac_url + '");><font face="Arial" size="2" color="#000064">' + ac_descripcion + '</font></a><a href=""><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar" alt="Cerrar vista" title="Cerrar vista" onclick=javascript:func_cerrarcarpeta("' + ac_codigo + '","' + ac_url + '"); align=right style="margin-right: 5px;"></a><a href=javascript:func_vercarpeta("' + ac_codigo + '");func_refrescarcarpeta("' + ac_codigo + '","' + ac_url + '");><img src="imagenes/refrescarcarpeta.png" alt="Actualizar vista" title="Actualizar vista" border=0 align=right style="margin-right: 2px;"></a></div>'
					).appendTo($objTabla.find('thead tr'));					
					func_vercarpeta(ac_codigo);					func_refrescarcarpeta(ac_codigo,ac_url);
										//sin link '<div id="divcarpeta' + ac_codigo + '" onmousedown=func_vercarpeta("' + ac_codigo + '"); style="text-align: left; width: 178px; height: 26px; background-image: url(imagenes/carpetaon.jpg);"><img src="imagenes/vacio.png" height=4 width=178><font face="Arial" size="2" color="#00529B">&nbsp;</font><font face="Arial" size="2" color="#000064">' + ac_descripcion + '</font><a href=javascript:func_vercarpeta("' + ac_codigo + '");func_refrescarcarpeta("' + ac_codigo + '","' + ac_url + '");><img src="imagenes/refrescarcarpeta.png" alt="Actualizar vista" title="Actualizar vista" border=0></a><a href=""><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar" alt="Cerrar vista" title="Cerrar vista" onclick=javascript:func_cerrarcarpeta("' + ac_codigo + '","' + ac_url + '");></a></div>'
					//con link '<div id="divcarpeta' + ac_codigo + '" onmousedown=func_vercarpeta("' + ac_codigo + '"); style="text-align: left; width: 178px; height: 26px; background-image: url(imagenes/carpetaon.jpg);"><img src="imagenes/vacio.png" height=4 width=178><font face="Arial" size="2" color="#00529B">&nbsp;</font><a style="text-decoration: none;" href=javascript:func_vercarpeta("' + ac_codigo + '","' + ac_url + '");><font face="Arial" size="2" color="#000064">' + ac_descripcion + '</font></a><a href=javascript:func_vercarpeta("' + ac_codigo + '");func_refrescarcarpeta("' + ac_codigo + '","' + ac_url + '");><img src="imagenes/refrescarcarpeta.png" alt="Actualizar vista" title="Actualizar vista" border=0></a><a href=""><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar" alt="Cerrar vista" title="Cerrar vista" onclick=javascript:func_cerrarcarpeta("' + ac_codigo + '","' + ac_url + '");></a></div>'
										//'<font face="Arial" size="2" color="#00529B">' + ac_descripcion + '</font><img id="carpeta' + ac_codigo + '" class="clsEliminar" src="imagenes/cerrarcarpeta.png" border=0>'					//'<table cellpadding=0 cellspacing=0 border=0 height=26 width=178><THEAD><TR background="imagenes/carpetaon.jpg"><td width=178><font face=Arial size=2 color=#00529B>'+ ac_descripcion + '</font><img src="imagenes/cerrarcarpeta.png" border=0 id="carpeta' + ac_codigo + '" class="clsEliminar"></td></TR></THEAD></table>'					
					//adjuntamos los td's de la columna al body de la tabla
					//$('<td>').html(
					//	'<input type="text" size="4">'
					//).appendTo($objTabla.find('tbody tr'));
					
					//cambiamos el atributo colspan del pie de la tabla y su contenido
					//$objTabla.find('tfoot tr td').attr('colspan',iTotalColumnasExistentes).
					//text('La tabla tiene '+iTotalColumnasExistentes+' columnas');				//}
				//else alert("Cierre alguna vista para poder continuar.");			}
			else
			{					func_ocultartotalcarpetas();
					func_vercarpeta(ac_codigo);			}
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
					
					/*					//Si la eliminada es visible entonces mostramos la primera que esté disponible en imagen $("#carpeta" + ac_codigo).length==0					//ojo que el resto ya está oculto
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
					*/					//preparamos el mensaje que vamos a mostrar en el pie de la tabla
					//var strMensaje='La tabla tiene '+iTotalColumnasExistentes+
					//((iTotalColumnasExistentes==1)?' columna':' columnas');
					//ajustamos el atributo colspan del pie de la tabla
					//$objTabla.find('tfoot tr td').attr('colspan',iTotalColumnasExistentes).text(strMensaje);				});		
				
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
				}
				function func_cerrarcarpeta(ac_codigo)
				{
					//func_vercarpeta(ac_codigo,ac_url);
					//alert(document.getElementById("fr_carpeta" + ac_codigo).style.visibility);
					if(document.getElementById("fr_carpeta" + ac_codigo).style.visibility=="visible")
					{
						//Si la eliminada es visible entonces mostramos la primera que esté disponible en imagen $("#carpeta" + ac_codigo).length==0						//ojo que el resto ya está oculto
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
						//Si la eliminada es visible entonces mostramos la primera que esté disponible en imagen $("#carpeta" + ac_codigo).length==0						//ojo que el resto ya está oculto
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
		</head>
		<body bgcolor="#FFFFFF" topmargin=0 leftmargin=0>
			<table width="100%" border=0 cellpadding=2 cellspacing=0>
				<tr>
					<td><img src="imagenes/logo.gif" alt="BBVA Continental" title="BBVA Continental"></td>
					<td valign=middle><font face="Arial" size=4 color="#00529B"><b>Sistema Web de Gestión de Cobranzas BBVA</b></font></td>
					<td align=right valign=bottom>
						<table border=0 cellpadding=2 cellspacing=0>
							<tr>
								<td align=right valign=middle><a href="javascript:olvideclave();" style="text-decoration:none;"><font face="Arial" size=2 color="#00529B">Modificar&nbsp;contraseña&nbsp;</font></a></td>					
								<td align=right valign=middle width="20"><a href="javascript:olvideclave();" style="text-decoration:none;"><img src="imagenes/modclave.png" alt="Modificar Contraseña" title="Modificar Contraseña" border=0></a></td>
								<td align=right valign=middle><font face="Arial" size=2 color="#00529B">&nbsp;|&nbsp;</font><a href="userexpira.asp" style="text-decoration:none;" target="_top"><font size=2 face="Arial" color="#00529B">&nbsp;Salir&nbsp;</font></a></td>
								<td align=right valign=middle><a href="userexpira.asp" style="text-decoration:none;" target="_top"><img src="imagenes/logout.png" alt="Salir" title="Salir" border=0></a><font size=2 face="Arial" color="#00529B">&nbsp;</font></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<form name="formula" method=post>
			<table width="99.6%" height="24" border=0 cellpadding=2 cellspacing=0 align=center>
				<tr>
					<td bgcolor="#BEE8FB"><font face="Arial" size=2 color="#00529B">&nbsp;Usuario: <b><%=session("nombreusuario")%></b></font></td>
					<td bgcolor="#BEE8FB" align=right valign=middle>
					<td bgcolor="#BEE8FB" align=right><font face="Arial" size=2 color="#00529B"><b>Perfil:</b>&nbsp;<select name="codperfil" style="font-size: xx-small; width: 200px;" onchange="document.formula.submit();">
								
								<%
									sql="select C.CodPerfil,C.Descripcion from Usuario A inner join UsuarioPerfil B on A.codusuario=B.codusuario inner join Perfil C on B.codperfil=C.codperfil where A.codusuario=" & session("codusuario") & " UNION select A.CodPerfil,A.Descripcion from Perfil A where (select administrador from Usuario where codusuario=" & session("codusuario") & ")=1 order by CodPerfil"
									consultar sql,RS
									codperfilselected=0
									Do while not RS.EOF 
										if obtener("codperfil")="" and codperfilselected=0 then
											codperfilselected=RS.Fields("CodPerfil")
										else
											if obtener("codperfil")<>"" then
												if int(obtener("codperfil"))=RS.Fields("CodPerfil") then
													codperfilselected=RS.Fields("CodPerfil")
												end if
											end if
										end if
									%>
										<option value=<%=RS.Fields("CodPerfil")%> <%if codperfilselected=RS.Fields("CodPerfil") then%> selected<%end if%>><%=RS.Fields("Descripcion")%></option>
									<%
									RS.MoveNext
									Loop
									RS.Close
								%>
								</select>&nbsp;</font></td>
				</tr>				
			</table>
			</form>
			<table width="100%" border=0 cellpadding=0 cellspacing=0>			
			<tr>
				<td>
					<!--menu inicio-->
					<div id="ddsidemenubar" class="markermenu">
					<ul>
					<li><a rel="ddsubmenuside1">Inicio</a></li>
					</ul>
					</div>
					 
					<script type="text/javascript"> 
					ddlevelsmenu.setup("ddsidemenubar", "sidebar") //ddlevelsmenu.setup("mainmenuid", "topbar|sidebar")
					</script>
					<!--fin menu inicio-->
				</td>
				<td align=right><font face="Arial" size=2 color="#00529B"><%fechalarga=FormatDateTime(Date(),1)%><%=UCase(mid(fechalarga,1,1))%><%=mid(fechalarga,2,len(fechalarga)-1)%>&nbsp;</font></td>
			</tr>
			
			
			<!--carpetas-->
			<TABLE id=tblTabla class=clsTabla cellspacing=0 cellpadding=0 border=0><THEAD><TR></TR></THEAD></TABLE>
			<!--fin carpetas-->


			<!--visibilidad de carpetas-->
			<IFRAME SRC="progvacio.html" style="visibility: hidden; position: absolute;" allowTransparency="true" HSPACE=0 ALIGN=TOP WIDTH=0 HEIGHT=0 FRAMEBORDER=0 NAME="fr_carpeta1" ID="fr_carpeta1" SCROLLING=AUTO><BLOCKQUOTE><P>Debe utilizar IExplorer 5.5 o superior.</P></BLOCKQUOTE></IFRAME>	
			<IFRAME SRC="progvacio.html" style="visibility: hidden; position: absolute;" allowTransparency="true" HSPACE=0 ALIGN=TOP WIDTH=0 HEIGHT=0 FRAMEBORDER=0 NAME="fr_carpeta2" ID="fr_carpeta2" SCROLLING=AUTO><BLOCKQUOTE><P>Debe utilizar IExplorer 5.5 o superior.</P></BLOCKQUOTE></IFRAME>
			<IFRAME SRC="progvacio.html" style="visibility: hidden; position: absolute;" allowTransparency="true" HSPACE=0 ALIGN=TOP WIDTH=0 HEIGHT=0 FRAMEBORDER=0 NAME="fr_carpeta3" ID="fr_carpeta3" SCROLLING=AUTO><BLOCKQUOTE><P>Debe utilizar IExplorer 5.5 o superior.</P></BLOCKQUOTE></IFRAME>				
			<!--fin visibildad carpetas-->

			<!--datos del menu:  Inicio-->
			<ul id="ddsubmenuside1" class="ddsubmenustyle blackwhite">
			<li><a>Mantenimiento</a>	
			  <ul>
			  <li><a href="javascript:agregarcol('1','Perfil','perfil.html');">Perfil</a></li>
			  <li><a href="javascript:agregarcol('2','Usuario','admusuario.asp');">Usuario</a></li>
			  <li><a href="javascript:agregarcol('3','Gestor','gestor.html');">Gestor</a></li>
			  <li><a href="javascript:agregarcol('4','Facultad','perfil.html');">xxxx</a></li>
			  <li><a href="javascript:agregarcol('5','Grupo Facultad','usuario.html');">yyyy</a></li>
			  <li><a href="javascript:agregarcol('6','Diseño','gestor.html');">zzzz</a></li>		
			  <li><a href="javascript:agregarcol('7','MantenimientoAAAAAddddddd','perfil.html');">aaaa</a></li>
			  <li><a href="javascript:agregarcol('8','MantenimientoBBBBBddddddd','usuario.html');">bbbb</a></li>
			  <li><a href="javascript:agregarcol('9','MantenimientoCCCCCdddddddd','gestor.html');">cccc</a></li>	
			  <li><a href="javascript:agregarcol('10','MantenimientoAAAAAddddddd','perfil.html');">aaaa</a></li>
			  <li><a href="javascript:agregarcol('11','MantenimientoBBBBBddddddd','usuario.html');">bbbb</a></li>
			  <li><a href="javascript:agregarcol('12','MantenimientoCCCCCdddddddd','gestor.html');">cccc</a></li>				  			  	  
			  <li><a href="#">Parámetro</a></li>
			  </ul>		
			</li>
			<li><a>Item 2a</a>
			  <ul>
			  <li><a href="#">Item 2a.xxx</a></li>
			  <li><a href="#">Item 2a.yyy</a></li>
			  <li><a href="#">Item 2a.zzz</a></li>
			  </ul>	
			</li>
			<li><a>Item 3</a>
			  <ul>
			  <li><a href="#">Item 2a.xxx</a></li>
			  <li><a href="#">Item 2a.yyy</a></li>
			  <li><a href="#">Item 2a.zzz</a></li>
			  <li><a href="#">Item 2a.zzz</a></li>
			  </ul>	
			</li>		
			<li><a>Item 2a</a>
			  <ul>
			  <li><a href="#">Item 2a.xxx</a></li>
			  <li><a href="#">Item 2a.yyy</a></li>
			  <li><a href="#">Item 2a.zzz</a></li>
			  </ul>	
			</li>		
			<!--fin datos del menu:  Inicio-->
		
		</body>
	</html>
<%
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



