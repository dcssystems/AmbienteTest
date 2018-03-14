<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
	conectar
	if permisofacultad("dcs_admfacultad.asp") then
		buscador=obtener("buscador")
		paginado=obtener("paginado")
		filtrobuscador2=obtener("filtrobuscador2")
		seltodo = obtener("seltodo")
		ordencampo = obtener("ordencampo")
		ordentipo = obtener("ordentipo")

		if ordentipo="" then
		ordentipo ="0"
		end if

		if ordencampo="" then
		ordencampo ="0"
		end if

		if seltodo = "" then
		seltodo = 0
		end if

		if paginado ="" then
		paginado = "18"
	    end if


			'response.write buscador2

		
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
			<script src="https://unpkg.com/sweetalert/dist/sweetalert.min.js"></script>

			<!--[if IE 7]><link rel="stylesheet" href="css/fontello-ie7.css"><![endif]-->	
			
			<script language="javascript" src="assets/jquery/dist/jquery-3.3.1.js"></script>

			<script language="javascript">
				$(document).ready(function(){
					$("#modal-filtro").hide();
					
				    $("#close-modal").on('click', function(){
						$("#modal-filtro").hide();
				    });
					$("#show-filtro").on('click', function(){
						$("#modal-filtro").removeClass('no-visible');
						$("#modal-filtro").show();
					});
					
				});
			</script>
	
    <script language="javascript">
      function toggleCodes(on) {
        var obj = document.getElementById('icons');
      
        if (on) {
          obj.className += ' codesOn';
        } else {
          obj.className = obj.className.replace(' codesOn', '');
        }
      }
      /*$document.scroll(function() {
 			 $(".title").toggleClass(newClass, $document.scrollTop() >= 5);
		});*/
      
    </script>   	
		<script language="javascript" src="scripts/TablaDinamica.js"></script>
		<script language="javascript">	

  	    var modfiltro = 0;

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
			ventanafacultad=global_popup_IWTSystem(ventanafacultad,"dcs_definiraccion.asp?vistapadre=" + window.name + "&paginapadre=dcs_definiraccion.asp","NewFacultad","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 180)/2 - 30) + ",height=180,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
			document.formula.buscador.value = document.formula.textobuscar.value;
			document.formula.pag.value=1;
			document.formula.submit();
		}
		function buscar2()
		{
			document.formula.pag.value=1;
			document.formula.buscador.value = "";
			document.formula.textobuscar.value = "";
			document.formula.submit();
		}			
		function habilitarorden()
		{
			if(document.getElementById("ordencampo").value != 0)
			{
				document.getElementById('ordentipo').disabled=false;
			}
			else
			{
				document.getElementById('ordentipo').value=0;
				document.getElementById('ordentipo').disabled=true;
			}
		}
		// function selectall(form)  
		// {  
 	// 		var formulario = eval(form)  
		//  for (var i=0, len=formulario.elements.length; i<len ; i++)  
		//   {  
		//     if ( formulario.elements[i].type == "checkbox" )  
		//       formulario.elements[i].checked = formulario.elements[0].checked  
		//   }  
		// }  
		
		function vercajatexto2(cajaid, valorcombo)	
		{				
			
			valorcombo =  document.getElementById(valorcombo).options[document.getElementById(valorcombo).selectedIndex].value
			  

			if (valorcombo == 14 || valorcombo == 12 || valorcombo == 8)
			{
			document.getElementById(cajaid).style.display = "inline";
			} 
			else
			{
			document.getElementById(cajaid).style.display = "none";
			}
			
		}

		function marcar(source) 
		{
			if (document.getElementById("seltodo").value == 1 )
		 	 {

		 	    document.getElementById("seltodo").value = 0;
		    }
		    else
		    {
		    	document.getElementById("seltodo").value = 1;
		    }

		 checkboxes=document.getElementsByTagName('input'); //obtenemos todos los controles del tipo Input
			for(i=0;i<checkboxes.length;i++) //recoremos todos los controles
			{
				if(checkboxes[i].type == "checkbox") //solo si es un checkbox entramos
				{
					checkboxes[i].checked=source.checked; //si es un checkbox le damos el valor del checkbox que lo llamÃ³ (Marcar/Desmarcar Todos)
				}
			}
		}
		// function seleccionar_todo(){ 	
			
		// 	 if (document.getElementById("seltodo").value == 1 )
		// 	 {

		// 	 	document.getElementById("seltodo").value = 0;
		// 	 	for (i=0;i<document.formula.elements.length;i++) 
		// 	 	{
		//      	 	if(document.formula.elements[i].type == "checkbox")	
		//      	 	{
		//        		document.formula.elements[i].checked=0;
		//        		}
	 //       		}
		// 	 }
		// 	 else
		// 	 {
		// 	 	document.getElementById("seltodo").value = 1;
		// 	 	for (i=0;i<document.formula.elements.length;i++) 
		// 	 	{
		//      	 	if(document.formula.elements[i].type == "checkbox")	
		//      	 	{
		//        		document.formula.elements[i].checked=1;
		//        		}
	 //       		}
		// 	 }

	  		 
	     	
		// } 
		// function deseleccionar_todo(){ 
  //  		for (i=0;i<document.f1.elements.length;i++) 
  //    	 if(document.f1.elements[i].type == "checkbox")	
  //        document.f1.elements[i].checked=0 
		// }
		
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
		function mostrarfiltro()
		{
			var filtro = document.getElementsByClassName('filtro-oculto').className="filtro-visible";
			console.log(filtro);
			//filtro.remove('visibility');			
		}

		


		</script>
		
		</head>
		
		<script language="javascript">
			
			<%
				idcampana = obtener("idcampana")''idcampana=2 	
				filtrobuscador = " where a.IDCampaña = " & idcampana 	



				' if  buscador2 <>""  then
				' 	filtrobuscador2 = filtrobuscador & " and a.IDCampañaPersona in ( select b.IDCampañaPersona from Campaña_Detalle a inner join Campaña_Persona b on a.IDCampañaPersona = b.IDCampañaPersona where b.IDCampaña = " & idcampana & " and ( "					
				' end if 

				if buscador<>""  then
					filtrobuscador = filtrobuscador & " and a.IDCampañaPersona in ( select b.IDCampañaPersona from Campaña_Detalle a inner join Campaña_Persona b on a.IDCampañaPersona = b.IDCampañaPersona where b.IDCampaña = " & idcampana & " and ( (b.NroDocumento like '%" & buscador & "%')"					
				end if 

				sql = "select ROW_NUMBER() OVER(ORDER BY IDCampañaCampo ) AS nro ,IDCampañaCampo, GlosaCampo, TipoCampo, FlagNroDocumento from Campaña_Campo a inner join Campaña b on a.IDTipoCampaña = b.IDTipoCampaña where a.Nivel = 1 and b.IDCampaña = " & IDCampana
							consultar sql,RS
							
						Do While Not RS.EOF

							if trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) <> "" Then

							' Response.Write "dato: " & obtener("buscador2" & RS.Fields("IDCampañaCampo")) & obtener("idTipocampo" & RS.Fields("IDCampañaCampo"))
							Select Case CInt(trim(obtener("idTipocampo" & RS.Fields("IDCampañaCampo"))))
									case 0
										  if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 13 then 
										  	filtrobuscadorX = " and (b.NroDocumento = '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  end if 
										  if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 2 then
										  	filtrobuscadorX = " and (b.NroDocumento like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										  end if 
										  if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 3 then
										  	filtrobuscadorX = " and (b.NroDocumento like '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										  end if 
										   if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 4 then
										  	filtrobuscadorX = " and (b.NroDocumento like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  end if 										    
										   if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 15 then
										  	filtrobuscadorX = " and (b.NroDocumento >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  end if
										  if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 16 then
										  	filtrobuscadorX = " and (b.NroDocumento <= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  end if
										  if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) =  14 then
										  	filtrobuscadorX = " and (b.NroDocumento >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo")))
										  	
										  	filtrobuscadorX = filtrobuscadorX & "' and b.NroDocumento <= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) & "' )"
										  end if
							
										  
							        case 1 
							        		if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampañaCampo =  " & trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto = '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto like '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 4 then
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto <= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"			  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorTexto >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo")))
										  
										  	filtrobuscadorX = filtrobuscadorX & "' and ValorTexto <='" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) & "')"
										    end if							             
							        case 2 
							             if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero = '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero like '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 4 then
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero <= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorEntero >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo")))
										  	
										  	filtrobuscadorX = filtrobuscadorX & "' and ValorEntero <='" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) & "')"
										    end if	
							        case 3 
							              if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat = '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat like '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 4 then
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat <= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFloat >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo")))
										  	
										  	filtrobuscadorX = filtrobuscadorX  & "' and ValorFloat <='" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) & "')"
										    end if		
							        case 4
							             if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha = '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha like '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 4 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha like '%" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha <= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "')"
										  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampañaCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampañaCampo =  "& trim(RS.Fields("IDCampañaCampo")) & " and ValorFecha >= '" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) 
										  	
										  	filtrobuscadorX = filtrobuscadorX & "' and ValorFecha <='" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) & "')"
										    end if
								End Select

								'response.Write "Dato: " & filtrobuscadorX

								if idpersonasfiltro = "" then
									sql="select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),a.IDCampañaPersona) from Campaña_Detalle a inner join Campaña_Persona b on a.IDCampañaPersona = b.IDCampañaPersona where b.IDCampaña = "& IDCampana & filtrobuscadorX & "  GROUP BY a.IDCampañaPersona ORDER BY a.IDCampañaPersona FOR XML PATH('') ), 1, 1, '') as Cadena"
									consultar sql,RS4
									else
									sql="select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),a.IDCampañaPersona) from Campaña_Detalle a inner join Campaña_Persona b on a.IDCampañaPersona = b.IDCampañaPersona where b.IDCampaña = "& IDCampana & filtrobuscadorX & " AND a.IDCampañaPersona in (" & idpersonasfiltro & " )  GROUP BY a.IDCampañaPersona ORDER BY a.IDCampañaPersona FOR XML PATH('') ), 1, 1, '') as Cadena"
									consultar sql,RS4
								end if


									if RS4.Fields("Cadena") <> "" then
										idpersonasfiltro = RS4.Fields("Cadena")										
										RS4.Close
									else
										%> 
										swal("No existen datos con ese filtro");
										<%
										idpersonasfiltro = ""
										RS4.Close
										exit do
									end if			


							end if

						RS.MoveNext
						loop
						RS.MoveFirst

						IF filtrobuscador2 ="" Then
						filtrobuscador2 = idpersonasfiltro
						end if

						

				sql="select GlosaCampo,ROW_NUMBER () over (order by NroCampo) as Orden,CampoCalculado,Formula,Condicion,IDCampañaCampo,TipoCampo,FlagNroDocumento,anchocolumna,aligncabecera,aligndetalle,alignpie,decimalesnumero,formatofecha " & chr(10) & _
                    "from Campaña_Campo " & chr(10) & _
                    "where IDTipoCampaña in (select IDTipoCampaña from Campaña where idcampaña=" & idcampana & ") " & chr(10) & _
                    "and Nivel=1 and Visible=1 " & chr(10) & _
                    "order by Orden"
                    'response.write "/*" & sql & "*/"
				consultar sql,RS3	
				nrocampos=RS3.RecordCount
				glosacampos=""
				glosavisible=""
				glosaancho=""
				glosaalineamiento=""
				Do while not RS3.EOF 
					Do While Not RS.EOF
					IF RS.fields("IDCampañaCampo")	= RS3.fields("IDCampañaCampo")	 Then
						if trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) <> "" and trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) <> ""  Then
						glosacampos=glosacampos & ",'" & RS3.Fields("GlosaCampo") & "<i class=" & chr(34) & "demo-icon3 icon-filter" & chr(34) & " title= "& chr(34) & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & "-" & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")) & chr(34) & ">&#xe820;</i>'"						
					    else
						    if trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) <> "" then
						   	 glosacampos=glosacampos & ",'" & RS3.Fields("GlosaCampo") & "<i class=" & chr(34) & "demo-icon3 icon-filter" & chr(34) & " title= "& chr(34) & trim(obtener("buscador2" & RS.Fields("IDCampañaCampo"))) & chr(34) & ">&#xe820;</i>'"
					    	else
					    	 glosacampos=glosacampos & ",'" & RS3.Fields("GlosaCampo") & "'"
					        end if
					    end if
					END IF
					
					RS.MoveNext
				    loop
				    RS.MoveFirst
					glosavisible=glosavisible & ",'true'"
					glosaancho=glosaancho & ",'" & RS3.Fields("anchocolumna") & "'"
					glosaaligncabecera=glosaaligncabecera & ",'" & RS3.Fields("aligncabecera") & "'"
					glosaaligndetalle=glosaaligndetalle & ",'" & RS3.Fields("aligndetalle") & "'"
					glosaalignpie=glosaalignpie & ",'" & RS3.Fields("alignpie") & "'"
					glosadecimalesnumero=glosadecimalesnumero & ",'" & RS3.Fields("decimalesnumero") & "'"
					glosaformatofecha=glosaformatofecha & ",'" & RS3.Fields("formatofecha") & "'"
					glosapie=glosapie & ",'&nbsp;'"
					glosapiefunciones=glosapiefunciones & ",''"								

					
					
					if buscador<>"" then					
					
					Select Case RS3.Fields("TipoCampo")
				        case 1 
				               filtrobuscador = filtrobuscador & " or (IDCampañaCampo =  "& RS3.Fields("IDCampañaCampo") & " and ValorTexto like '%" & buscador & "%')"
				        case 2 
				               filtrobuscador = filtrobuscador & " or (IDCampañaCampo =  " & RS3.Fields("IDCampañaCampo") & " and ValorEntero like '%" & buscador & "%')"
				        case 3 
				               filtrobuscador = filtrobuscador & "  or (IDCampañaCampo =  "& RS3.Fields("IDCampañaCampo") & " and ValorFloat like '%" & buscador & "%')"
				        case 4
				               filtrobuscador = filtrobuscador & " or (IDCampañaCampo =  "& RS3.Fields("IDCampañaCampo") & " and ValorFecha like '%" & buscador & "%')"
					End Select
		
					end if
				RS3.MoveNext 
				Loop
				RS.Close
				RS3.MoveFirst

				if buscador<>"" then
					filtrobuscador = filtrobuscador & "))"
					'response.write filtrobuscador
				end if 
				' if buscador2<>"" then
				' 	filtrobuscador2 = filtrobuscador2 & "))"
				' end if
				 

			%>

			
			rutaimgcab="imagenes/"; 
		  //ConfiguraciÃ³n general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=0;
		    ascendente[tabla]=true;
		    nrocolumnas[tabla]=<%=nrocampos + 1%>;
		    fondovariable[tabla]='bgcolor=#e9f7f7';
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
		    anchocolumna[tabla] =  new Array(''<%=glosaancho%>);
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
		
		'if buscador<>"" then
		''			filtrobuscador = "where a.IDCampaña =" & idcampana & " and a.IDCampañaPersona in ( select b.IDCampañaPersona from Campaña_Detalle a inner join Campaña_Persona b on a.IDCampañaPersona = b.IDCampañaPersona where b.IDCampaña = " & idcampana & " and IDCampañaCampo = 1 and ValorTexto like '%" & buscador & "%')"
		''		end if'
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if		
				
		
		contadortotal=0
		
		'sql="select Count(*) / (select count(*) from Campaña_Campo b inner join campaña c on b.IDTipoCampaña = 'c.IDTipoCampaña where c.IDCampaña =" & idcampana & " and b.FlagNroDocumento <> 1) from Campaña_Detalle a where 'IDCampañaPersona in (Select IDCampañaPersona from Campaña_Persona a " & filtrobuscador & " ) "  
		
		
		if trim(filtrobuscador2) <> "" then
		sql = "Select count(distinct IDCampañaPersona) from Campaña_Persona a " & filtrobuscador & " and a.IDCampañaPersona IN (" & filtrobuscador2 & " )"
		else
		sql = "Select count(distinct IDCampañaPersona) from Campaña_Persona a " & filtrobuscador 
		end if
		'Response.Write filtrobuscador2
		'' response.write sql
		consultar sql,RS	
		contadortotal=rs.fields(0)
		
		RS.Close
		
		if paginado <> "" then
		cantidadxpagina=paginado
	    else
	    cantidadxpagina=18
		end if
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
			sql="SELECT TOP " & cantidadxpagina & " A.IDCampañaPersona,A.NroDocumento from Campaña_Persona A where A.idcampaña=" & IDCampana & " and " & filtrobuscador1 & " A.IDCampañaPersona NOT  IN (SELECT TOP " & topnovisible & " A.IDCampañaPersona FROM Campaña_Persona A  " & filtrobuscador & " order by A.IDCampañaPersona) order by A.IDCampañaPersona"		
			else
			sql="SELECT TOP " & cantidadxpagina & " A.IDCampañaPersona,A.NroDocumento from Campaña_Persona A " & filtrobuscador & " order by A.IDCampañaPersona" 
			end if
			
			if filtrobuscador2 <> "" or (filtrobuscador2 <>"" and buscador <>"") then
				if pag>1 then			

				sql="SELECT TOP " & cantidadxpagina & " A.IDCampañaPersona,A.NroDocumento from Campaña_Persona A where " & filtrobuscador1 & " A.IDCampañaPersona NOT  IN (SELECT TOP " & topnovisible & " A.IDCampañaPersona FROM Campaña_Persona A  " & filtrobuscador & " and A.IDCampañaPersona in (" & filtrobuscador2 & ") ) and A.IDCampañaPersona in (" & filtrobuscador2 & ") order by A.IDCampañaPersona"
				else
				sql="SELECT TOP " & cantidadxpagina & " A.IDCampañaPersona,A.NroDocumento from Campaña_Persona A " & filtrobuscador & " and A.IDCampañaPersona in (" & filtrobuscador2 & ") order by A.IDCampañaPersona"	
				end if
			end if
	
		

		consultar sql,RS
		contador=0
 
		
		if RS.RecordCount > 0 then
			idpersonas =""
			Do while not RS.EOF 
				if idpersonas = "" then
					idpersonas = RS.Fields("IDCampañaPersona")
				Else
					idpersonas = idpersonas & "," & RS.Fields("IDCampañaPersona")
				end if				
			RS.MoveNext 
			Loop 	
			RS.MoveFirst
	




   		sql="select A.IDCampañaPersona,A.NroDocumento,C.IDCampañaCampo,C.NroCampo,C.TipoCampo,D.ValorTexto,D.ValorEntero,D.ValorFloat,D.ValorFecha " & chr(10) & _
            "from Campaña_Persona A " & chr(10) & _
            "inner join Campaña B " & chr(10) & _
            "on A.IDCampaña=B.IDCampaña " & chr(10) & _
            "inner join Campaña_Campo C " & chr(10) & _
            "on B.IDTipoCampaña=C.IDTipoCampaña and C.Nivel=1 and C.Visible=1 " & chr(10) & _
            "inner join Campaña_Detalle D " & chr(10) & _
            "on A.IDCampañaPersona=D.IDCampañaPersona and C.IDCampañaCampo=D.IDCampañaCampo " & chr(10) & _
            " where A.IDCampañaPersona IN (" & idpersonas & ")"
            'response.write sql
		consultar sql,RS2

		
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

			end if
		
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
		
		<%
				
		sql= "select Descripcion from campaña where IDCampaña = " & IDCampana
		consultar sql,RS
		descripcioncampana = rs.fields("Descripcion")
		RS.Close

		if contador=0 then%>
		
		<body topmargin="0" leftmargin="0" style="overflow-x:hidden;">
			<form name="formula" method="post">
				<table width="100%" cellpadding="4" cellspacing="0">	
					<tr class="fondo-orange">
						<td class="text-orange"><font size="2" face="Raleway"><b><%=descripcioncampana%> (0) - No hay registros.</b></font>&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-doc">&#xe808;</i></a></td>
						<td class="text-orange" align="middle" width="250"><font size="2" face="Raleway">Buscar:&nbsp;<input name="buscador" value="<%=buscador%>" size="20" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
						<td class="text-orange" align="left"><a href="javascript:buscar();"><i class="demo-icon icon-search">&#xe80c;</i></a></td>
					</tr>
				</table>
		<%else		
		%>		
		<body topmargin="0" leftmargin="0" style="overflow-x:hidden;"><!--onload="inicio();"-->
			<div id="modal-filtro" class="filtro-visible no-visible" >		
				<form name="formula" id="formula" method="post">		
					<table border="0">
						<tr class="fondo-red">
							<td class="text-withe" colspan="3">
								Realizar Filtro							
								<a style="float:right; padding-right:0px;" href="javascript:limpiarform();"><i style="color: white;"  class="demo-icon2 icon-reply">&#xe81e;</i></a>
							</td>
							<td id="close-modal"><a style="float:right; padding-right:0px;" href="javascript:if (modfiltro != 0 ){ buscar2();}"><i style="color: white;" class="demo-icon2 icon-cancel-circle">&#xe807;</i></a></td>
						</tr>
						<tr class="fondo-red" >
							<td class="text-withe">Campo</td>
							<td class="text-withe">Filtro</td>
							<td class="text-withe">Dato</td>							
						</tr>
						<%
							sql = "select ROW_NUMBER() OVER(ORDER BY IDCampañaCampo ) AS nro ,IDCampañaCampo, GlosaCampo, TipoCampo, FlagNroDocumento from Campaña_Campo a inner join Campaña b on a.IDTipoCampaña = b.IDTipoCampaña where a.Nivel = 1 and b.IDCampaña = " & IDCampana
							consultar sql,RS
							
							cadenareset = ""
							Do While Not RS.EOF								
						%>

						<tr class="fondo-red <% IF(CInt(RS.Fields("nro")) mod 2) <> 0 Then %> fondo-blanco <% Else %> fondo-rojo <% End IF %>" >
							<input type="hidden" name="idcampanacampo<%=RS.fields("IDCampañaCampo")%>" value="<%=obtener("idcampanacampo" & RS.fields("IDCampañaCampo"))%>">							
							<input type="hidden" name="idTipocampo<%=RS.Fields("IDCampañaCampo")%>" value="<%If RS.Fields("FlagNroDocumento") = 0 then%><%=RS.Fields("TipoCampo")%><%Else%>0<%End if%>">
							<td>							
							<%=RS.Fields("GlosaCampo")%>
							</td>

							<td class="text-withe" width="120" style="text-align: center;">
								<%cadenareset = cadenareset & "document.formula.filtrado" & RS.fields("IDCampañaCampo") & ".value=0;" & chr(13)%>
									<select name="filtrado<%=RS.fields("IDCampañaCampo")%>" id="Select<%=RS.Fields("IDCampañaCampo")%>" onChange="javascript:modfiltro++;vercajatexto2('<%=trim(RS.fields("IDCampañaCampo")) & "b"%>','Select<%=RS.Fields("IDCampañaCampo")%>');" style="font-size: xx-small; text-align: center; width: 100px;">
										<option value="0">Seleccione un filtro</option>
										<%
										sql = "SELECT idfiltro, descripcion FROM Filtro WHERE TipoCampo =" & RS.Fields("TipoCampo") 
										consultar sql,RS4
										Do While Not  RS4.EOF
										%>					
										
											<option value="<%=RS4.Fields("idfiltro")%>" 
												<% if obtener("filtrado" & RS.fields("IDCampañaCampo")) <> "" Then
												 if RS4.fields("idfiltro") = CInt(obtener("filtrado" & RS.fields("IDCampañaCampo"))) Then %> 
												 selected
												<%end if 
												end if%> 
											    ><%=RS4.Fields("descripcion")%>								    	
											 </option>									
												
										<%										
										RS4.MoveNext
										loop
										RS4.Close
										%>
									</select>
							</td>
							<td class="text-orange">
								<%cadenareset = cadenareset & "document.formula.buscador2" & RS.fields("IDCampañaCampo") & ".value='';" & chr(13)%>
								<input type="text" id="<%=RS.Fields("IDCampañaCampo")%>" onChange="javascript:modfiltro++;" name="buscador2<%=RS.Fields("IDCampañaCampo")%>"  class="form-control" value="<%=obtener("buscador2" & RS.Fields("IDCampañaCampo"))%>"/>
								<% 
								if RS.Fields("TipoCampo") = "2" or RS.Fields("TipoCampo") ="3" or RS.Fields("TipoCampo") = "4" then 									
								%>

								<%cadenareset = cadenareset & "document.formula.buscador2" & RS.fields("IDCampañaCampo") & "b.value='';" & chr(13)%>
								<input type="text"  onChange="javascript:modfiltro++;" id="<%=RS.Fields("IDCampañaCampo")%>b"
								<% if obtener("filtrado" & RS.fields("IDCampañaCampo")) <> "" Then 
								if CInt(obtener("filtrado" & RS.fields("IDCampañaCampo"))) <> 12 and CInt(obtener("filtrado" & RS.fields("IDCampañaCampo"))) <> 8 and CInt(obtener("filtrado" & RS.fields("IDCampañaCampo"))) <> 14 Then 
								%>
								 style="display:none;" <% end if 
								 else %>
								  style="display:none;"
								 <%end if %>name="buscador2<%=RS.Fields("IDCampañaCampo")%>b"  class="form-control" value="<%=obtener("buscador2" & RS.Fields("IDCampañaCampo") & "b")%>"/>
								 <% end if %>
							</td>							
						</tr>
						<%
							RS.MoveNext
							loop
							RS.MoveFirst
							%>
					</table>			
					<script type="text/javascript">
						function limpiarform()
						{
							<%=cadenareset%>
							modfiltro++;
						}
					</script>			
				</div>						
				<table width="100%" cellpadding="4" cellspacing="0" border="0"><!--Esto no sale -->	
					<tr class="fondo-orange">
						<td class="text-orange" align="left" width="250"><font size="2" face="Raleway"><b><%=descripcioncampana%> (<%=contadortotal%>)</td>
							<td class="text-orange" align="right" width="250"><font size="2" face="Raleway">Buscar:&nbsp;<input name="textobuscar" value="<%=buscador%>" size="20" id="textobuscar" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
							<td width="80">
								<a href="javascript:buscar();"><i class="demo-icon icon-search">&#xe80c;</i></a>
								<a id="show-filtro" href="#"><i class="demo-icon icon-filter">&#xe820;</i></a>							
							</td>
							<td width="125">Ver:
								<SELECT name="paginado" onChange="javascript:buscar();">
								<option value="18" <%if CInt(paginado) = 18 then %>selected<%end if%> >18</option>
								<option value="50" <%if CInt(paginado) = 50 then %>selected<%end if%> >50</option>
								<option value="100" <%if CInt(paginado) = 100 then %>selected<%end if%>  >100</option>
								<option value="500" <%if CInt(paginado) = 500 then %>selected<%end if%> >500</option>
								<option value="1000" <%if CInt(paginado) = 1000 then %>selected<%end if%> >1000</option>
								<!-- <option value="<%=contadortotal%>" <%if CInt(paginado) = contadortotal then %>selected<%end if%> >Todos</option> -->
								</select>					
							</td>
							<td>
								<Select name="ordencampo"  id="ordencampo" onChange="javascript:habilitarorden();" >
									<option value="0">Seleccione Campo Orden</option>
								<%	Do While Not RS.EOF								
								%>
									<option value="<%=RS.Fields("IDCampañaCampo")%>" <%if CInt(ordencampo) =  RS.Fields("IDCampañaCampo") then %>selected<%end if%>><%=RS.Fields("GlosaCampo")%></option>
								<%
									RS.MoveNext
									loop
									RS.MoveFirst
								%>
								</Select>
								<Select name="ordentipo" id="ordentipo" onChange="javascript:buscar();" <%if Cint(ordencampo)=0 then%> disabled <%end if%>>
									<option value="0">Seleccione vector</option>
									<option value="asc" <%if Cint(ordentipo) = 1 then %> selected <%end if%>>Ascendente</option>
									<option value="desc" <%if Cint(ordentipo) = 2 then %> selected <%end if%>>Descendente</option>
								</Select>
							</td>

						<td class="text-orange" align="right">
							&nbsp;&nbsp;<a href="javascript:agregar();"><i class="demo-icon icon-cog">&#xe81f;</i></a>&nbsp;&nbsp;<a href="javascript:exportar();"><i class="demo-icon icon-file-excel">&#xf1c3;</i></a>
						<%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><i class="demo-icon icon-download">&#xe814;</i></a><%end if%></b></font>
						</td>
						<td class="text-orange" align="right" width="180"><font size="2" face="Raleway">PÃ¡g.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
					</tr>	
				</table>
				<div id="tabla0"> 
				</div>
		<%end if%>
		
		<input type="hidden" name="actualizarlista" value="">
		<input type="hidden" name="expimp" value="">		
		<input type="hidden" name="buscador" value="<%=buscador%>">		
		
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
	

		
	

		
	else
	%>
	<script language="javascript">
		alert("Ud. No tiene autorizaciÃ³n para este proceso.");
		window.open("dcs_userexpira.asp","_top");
	</script>
	<%	
	
	end if	
	
	if contador > 0 then
	RS3.Close
	RS2.Close
	else
	RS.Close
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



