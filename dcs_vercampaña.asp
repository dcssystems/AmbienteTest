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
		codusuario = obtener("codusuario")
		idpersonas_asig = obtener("idpersonas_asig")
		filoperador = obtener("filoperador")
		checkvisible = obtener("checkvisible")
		idpersonas_chk = obtener("idpersonas_chk")
		codrespuesta = obtener("codrespuesta")

	



		if ordencampo ="" then
			ordencampo = 0
		else
		ordencampo = CInt(ordencampo)
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
				$(document).ready(function(){
					$("#modal-filtro2").hide();
					
				    $("#close-modal2").on('click', function(){
						$("#modal-filtro2").hide();
				    });
					$("#show-filtro2").on('click', function(){
						$("#modal-filtro2").removeClass('no-visible');
						$("#modal-filtro2").show();
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
		function agregar(idcampana, personasasignar)
		{
			ventanafacultad=global_popup_IWTSystem(ventanafacultad,"dcs_definircall.asp?idcampana="+idcampana+"&personasasignar="+personasasignar+"&vistapadre=" + window.name + "&paginapadre=dcs_definiraccion.asp","NewFacultad","scrollbars=yes,scrolling=yes,top=" + ((screen.height - 180)/2 - 30) + ",height=180,width=" + (screen.width/2 - 10) + ",left=" + (screen.width/4) + ",resizable=yes");
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
			if(document.getElementById("ordencampo").value != "")
			{
				document.getElementById('ordentipo').disabled=false;
			}
			else
			{
				document.getElementById('ordentipo').value="";
				document.getElementById('ordentipo').disabled=true;
			}
		}

		
		function asignar(cruces)
		{

				if(formula.codusuario.value=="0"){swal("Debe escoger a un Operador");return;}

				document.formula.pag.value=1;

				

			if(cruces == 0)
			{			
				document.formula.idpersonas_asig.value = document.formula.idpersonas_chk.value;
				document.formula.submit();
			}
			else
			{
						
						if (cruces != "0")
						{
								swal("Existen " + cruces + " registros ya asignados, estas seguro de continuar con la asignaci�n.", {
						  dangerMode: true,
						   buttons: ["NO", "SI"],
						}).then(
					       function (isConfirm) {
								if (isConfirm) {
					           
					                document.formula.submit();
					           

					        } else {
					            swal("Cancelado", "no se realizo ninguna asignaci�n", "error");
					        }

					       });
					}
						else
						{

						  swal("Estas seguro de continuar con la asignaci�n.", {
						  dangerMode: true,
						   buttons: ["NO", "SI"],
						}).then(
					       function (isConfirm) {
								if (isConfirm) {
					           
					                document.formula.submit();
					           

					        } else {
					            swal("Cancelado", "no se realizo ninguna asignaci�n", "error");
					        }

					       });			
						}

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
				if(checkboxes[i].type == "checkbox" && checkboxes[i].name != "checkvisible" && checkboxes[i].name != "seltodo")  //solo si es un checkbox entramos
				{
					checkboxes[i].checked=source.checked;
					if 	(source.checked == true )
					{				
					checkarpersonas(checkboxes[i].name); //si es un checkbox le damos el valor del checkbox que lo llam� (Marcar/Desmarcar Todos)
					}
					else
					{
					document.formula.idpersonas_chk.value = "";
					}
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

		function activarchecks(chk)
		{
			if (chk.checked == true)
			{
				document.getElementById("checkvisible").value = "checked";
			}
			else
			{
				document.getElementById("checkvisible").value = "";
			}
			document.formula.submit();

		}

		function checkarpersonas(idpersona)
		{
			if(document.getElementById(""+idpersona+"").checked == true)
			{
				if (document.getElementById("idpersonas_chk").value == "")
				{
				document.formula.idpersonas_chk.value = idpersona;
				}
				else
				{
				document.formula.idpersonas_chk.value = document.formula.idpersonas_chk.value  + "," + idpersona;
				}
			}
			else
			{
				document.formula.idpersonas_chk.value  = document.formula.idpersonas_chk.value.replace(","+ idpersona + ",",",")
				document.formula.idpersonas_chk.value  = document.formula.idpersonas_chk.value.replace(","+ idpersona,"")
				document.formula.idpersonas_chk.value  = document.formula.idpersonas_chk.value.replace(""+ idpersona +",","")
				document.formula.idpersonas_chk.value  = document.formula.idpersonas_chk.value.replace(",,","")
			}
		}

		


		</script>
		
		</head>
		
		<script language="javascript">
			
			<%
				idcampana = obtener("idcampana")''idcampana=2 

				if ( filoperador = "") then					
				filtrobuscador = " where a.IDCampa�a = " & idcampana 
				else
					if ( filoperador = 0) then	
						filtrobuscador = " where a.IDCampa�a = " & idcampana & " and a.UsuarioAsignado is NULL " 
					else
						filtrobuscador = " where a.IDCampa�a = " & idcampana & " and a.UsuarioAsignado = " & filoperador
					end if
				end if 

				sql="Select Descripcion, convert(varchar(10),FechaInicio,103) as Inicio,  convert(varchar(10),fechafin,103) as Fin from Campa�a where idcampa�a =" & idcampana



				consultar sql,RS

				 Nombcampana=RS.Fields(0)
				 fechainicio=RS.Fields(1)
				 fechafin=RS.Fields(2)

				 RS.Close

				 if idpersonas_asig <> "" and codusuario <> "" then
							sql="update Campa�a_Persona set UsuarioAsignado = " & codusuario & " where IDCampa�aPersona in ( " & idpersonas_asig &" )"
							
							'response.write sql

							conn.Execute sql
							%>							
								swal("Se asigno correctamente",{icon: "success",  buttons: false,  timer: 3000,});							
							<%
							idpersonas_asig =""
							codusuario =""
				 end if
								

				' if  buscador2 <>""  then
				' 	filtrobuscador2 = filtrobuscador & " and a.IDCampa�aPersona in ( select b.IDCampa�aPersona from Campa�a_Detalle a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where b.IDCampa�a = " & idcampana & " and ( "					
				' end if 

				if buscador<>""  then
					filtrobuscador = filtrobuscador & " and a.IDCampa�aPersona in ( select b.IDCampa�aPersona from Campa�a_Detalle a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where b.IDCampa�a = " & idcampana & " and ( (b.NroDocumento like '%" & buscador & "%')"					
				end if 

				sql = "select ROW_NUMBER () over (order by NroCampo) AS nro ,IDCampa�aCampo, GlosaCampo, TipoCampo, FlagNroDocumento,CampoCalculado,Formula,Condicion,anchocolumna,aligncabecera,aligndetalle,alignpie,decimalesnumero,formatofecha,visible from Campa�a_Campo a inner join Campa�a b on a.IDTipoCampa�a = b.IDTipoCampa�a where a.Nivel = 1 and b.IDCampa�a = " & IDCampana & " order by nro"
							
								'response.Write sql

							consultar sql,RS
							
			

						Do While Not RS.EOF

							if trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) <> "" Then

							' Response.Write "dato: " & obtener("buscador2" & RS.Fields("IDCampa�aCampo")) & obtener("idTipocampo" & RS.Fields("IDCampa�aCampo"))
							Select Case CInt(trim(obtener("idTipocampo" & RS.Fields("IDCampa�aCampo"))))
									case 0
										  if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 13 then 
										  	filtrobuscadorX = " and (b.NroDocumento = '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  end if 
										  if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 2 then
										  	filtrobuscadorX = " and (b.NroDocumento like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										  end if 
										  if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 3 then
										  	filtrobuscadorX = " and (b.NroDocumento like '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										  end if 
										   if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 4 then
										  	filtrobuscadorX = " and (b.NroDocumento like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  end if 										    
										   if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 15 then
										  	filtrobuscadorX = " and (b.NroDocumento >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  end if
										  if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 16 then
										  	filtrobuscadorX = " and (b.NroDocumento <= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  end if
										  if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) =  14 then
										  	filtrobuscadorX = " and (b.NroDocumento >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo")))
										  	
										  	filtrobuscadorX = filtrobuscadorX & "' and b.NroDocumento <= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) & "' )"
										  end if
							
										  
							        case 1 
							        		if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampa�aCampo =  " & trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto = '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto like '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 4 then
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto <= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"			  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorTexto >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo")))
										  
										  	filtrobuscadorX = filtrobuscadorX & "' and ValorTexto <='" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) & "')"
										    end if							             
							        case 2 
							             if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero = '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero like '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 4 then
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero <= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorEntero >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo")))
										  	
										  	filtrobuscadorX = filtrobuscadorX & "' and ValorEntero <='" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) & "')"
										    end if	
							        case 3 
							              if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat = '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat like '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 4 then
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat <= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFloat >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo")))
										  	
										  	filtrobuscadorX = filtrobuscadorX  & "' and ValorFloat <='" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) & "')"
										    end if		
							        case 4
							             if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 1 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 5 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 9 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 13 then 

							        			filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha = '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
											end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 2 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										  	if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 3 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha like '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "%')"
										 	end if 
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 4 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha like '%" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if 										    
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 6 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 10 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 15 then 
										  	 filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										 	end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 7 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 11 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 16 then 
										     filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha <= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "')"
										  	
										    end if
										    if trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 8 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 12 or trim(obtener("filtrado" & RS.Fields("IDCampa�aCampo"))) = 14 then 
										  	filtrobuscadorX = " and (IDCampa�aCampo =  "& trim(RS.Fields("IDCampa�aCampo")) & " and ValorFecha >= '" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) 
										  	
										  	filtrobuscadorX = filtrobuscadorX & "' and ValorFecha <='" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) & "')"
										    end if
								End Select

								'response.Write "Dato: " & filtrobuscadorX

								if idpersonasfiltro = "" then
									sql="select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),a.IDCampa�aPersona) from Campa�a_Detalle a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where b.IDCampa�a = "& IDCampana & filtrobuscadorX & "  GROUP BY a.IDCampa�aPersona ORDER BY a.IDCampa�aPersona FOR XML PATH('') ), 1, 1, '') as Cadena"
									consultar sql,RS4
									else
									sql="select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),a.IDCampa�aPersona) from Campa�a_Detalle a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where b.IDCampa�a = "& IDCampana & filtrobuscadorX & " AND a.IDCampa�aPersona in (" & idpersonasfiltro & " )  GROUP BY a.IDCampa�aPersona ORDER BY a.IDCampa�aPersona FOR XML PATH('') ), 1, 1, '') as Cadena"
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



						if codrespuesta <> "" AND idpersonasfiltro <> ""  then
							sql = "select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),a.IDCampa�aPersona) " & chr(10) & _
							"from Campa�a_Persona A " & chr(10) & _
							"where a.IDCampa�aPersona in ( " & idpersonasfiltro & " ) and (select top 1 c.idGestion from Campa�a_Persona_Accion b " & chr(10) & _
							"inner join Gestion c on b.IDGestion = c.IDGestion " & chr(10) & _
							"where IDCampa�aPersona = A.IDCampa�aPersona " & chr(10) & _
							"order by prioridad asc , b.FechaRegistra desc) = " & codrespuesta  & chr(10) & _
							"GROUP BY a.IDCampa�aPersona ORDER BY a.IDCampa�aPersona FOR XML PATH('') ), 1, 1, '') as Cadena"

							
							consultar sql,RS4

							if RS4.Fields("Cadena") <> "" then
										idpersonasfiltro = RS4.Fields("Cadena")										
										RS4.Close
									else
										%> 
										swal("No existen datos con ese filtro");
										<%
										idpersonasfiltro = ""
										RS4.Close
									end if	
						Else

							if codrespuesta <> "" then
								sql = "select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),a.IDCampa�aPersona) " & chr(10) & _
								"from Campa�a_Persona A " & chr(10) & _
								"where a.IDCampa�a = " & idcampana & " and (select top 1 c.idGestion from Campa�a_Persona_Accion b " & chr(10) & _
								"inner join Gestion c on b.IDGestion = c.IDGestion " & chr(10) & _
								"where IDCampa�aPersona = A.IDCampa�aPersona " & chr(10) & _
								"order by prioridad asc , b.FechaRegistra desc) = " & codrespuesta  & chr(10) & _
								" GROUP BY a.IDCampa�aPersona ORDER BY a.IDCampa�aPersona FOR XML PATH('') ), 1, 1, '') as Cadena"

								
								consultar sql,RS4

								if RS4.Fields("Cadena") <> "" then
											idpersonasfiltro = RS4.Fields("Cadena")										
											RS4.Close
										else
											%> 
											swal("No existen datos con ese filtro");
											<%
											idpersonasfiltro = ""
											RS4.Close
										end if	
							end if


						end if


						IF filtrobuscador2 ="" Then
						filtrobuscador2 = idpersonasfiltro
						end if

						

				''sql="select GlosaCampo,ROW_NUMBER () over (order by NroCampo) as Orden,CampoCalculado,Formula,Condicion,IDCampa�aCampo,TipoCampo,FlagNroDocumento,anchocolumna,aligncabecera,aligndetalle,alignpie,decimalesnumero,formatofecha " & chr(10) & _
                    ' "from Campa�a_Campo " & chr(10) & _
                    ' "where IDTipoCampa�a in (select IDTipoCampa�a from Campa�a where idcampa�a=" & idcampana & ") " & chr(10) & _
                    ' "and Nivel=1 and Visible=1 " & chr(10) & _
                    ' "order by Orden"
                    'response.write "/*" & sql & "*/"
                    'response.Write sql
				''consultar sql,RS3	
				RS.Filter=" visible=1 "


				nrocampos=RS.RecordCount
				glosacampos=""
				glosavisible=""
				glosaancho=""				
				glosaalineamiento=""
				camposel=1
				mitablaorden=0
				Do while not RS.EOF 			
					camposel= camposel+1
					if RS.Fields("IDCampa�aCampo") = ordencampo then
					mitablaorden = camposel
					end if
					
						if trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) <> "" and trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) <> ""  Then
						glosacampos=glosacampos & ",'" & RS.Fields("GlosaCampo") & "<i class=" & chr(34) & "demo-icon3 icon-filter" & chr(34) & " title= "& chr(34) & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & "-" & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")) & chr(34) & ">&#xe820;</i>'"						
					    else
						    if trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) <> "" then
						   	 glosacampos=glosacampos & ",'" & RS.Fields("GlosaCampo") & "<i class=" & chr(34) & "demo-icon3 icon-filter" & chr(34) & " title= "& chr(34) & trim(obtener("buscador2" & RS.Fields("IDCampa�aCampo"))) & chr(34) & ">&#xe820;</i>'"
					    	else
					    	 glosacampos=glosacampos & ",'" & RS.Fields("GlosaCampo") & "'"
					        end if
					    end if





					glosavisible=glosavisible & ",'true'"
					glosaancho=glosaancho & ",'" & RS.Fields("anchocolumna") & "'"
					glosaaligncabecera=glosaaligncabecera & ",'" & RS.Fields("aligncabecera") & "'"
					glosaaligndetalle=glosaaligndetalle & ",'" & RS.Fields("aligndetalle") & "'"
					glosaalignpie=glosaalignpie & ",'" & RS.Fields("alignpie") & "'"
					glosadecimalesnumero=glosadecimalesnumero & ",'" & RS.Fields("decimalesnumero") & "'"
					glosaformatofecha=glosaformatofecha & ",'" & RS.Fields("formatofecha") & "'"
					glosapie=glosapie & ",'&nbsp;'"
					glosapiefunciones=glosapiefunciones & ",''"								

					
					
					if buscador<>"" then					
					
					Select Case RS.Fields("TipoCampo")
				        case 1 
				               filtrobuscador = filtrobuscador & " or (IDCampa�aCampo =  "& RS.Fields("IDCampa�aCampo") & " and ValorTexto like '%" & buscador & "%')"
				        case 2 
				               filtrobuscador = filtrobuscador & " or (IDCampa�aCampo =  " & RS.Fields("IDCampa�aCampo") & " and ValorEntero like '%" & buscador & "%')"
				        case 3 
				               filtrobuscador = filtrobuscador & "  or (IDCampa�aCampo =  "& RS.Fields("IDCampa�aCampo") & " and ValorFloat like '%" & buscador & "%')"
				        case 4
				               filtrobuscador = filtrobuscador & " or (IDCampa�aCampo =  "& RS.Fields("IDCampa�aCampo") & " and ValorFecha like '%" & buscador & "%')"
					End Select
		
					end if
				RS.MoveNext 
				Loop
				RS.MoveFirst

				

				if buscador<>"" then
					filtrobuscador = filtrobuscador & "))"
					'response.write filtrobuscador
				end if 
				' if buscador2<>"" then
				' 	filtrobuscador2 = filtrobuscador2 & "))"
				' end if
				 

			%>

			
			rutaimgcab="imagenes/"; 
		  //Configuraci�n general de datos de tabla 0
		    tabla=0;
		    orden[tabla]=<%=mitablaorden%>;
		    ascendente[tabla]=<% if ordentipo = "" or ordentipo = "asc" then %>true<%else%>false<%end if%>;
		    nrocolumnas[tabla]=<%=nrocampos + 6%>;
		    fondovariable[tabla]='bgcolor=#e9f7f7';
		    anchotabla[tabla]='100%';
		    botonfiltro[tabla] = false;
		    botonactualizar[tabla] = false;
		    botonagregar[tabla] = false;
			paddingtabla[tabla] = '0';
			spacingtabla[tabla] = '1';			    
		    cabecera[tabla] = new Array('IDCampanaPersona','Slc','Asignacion'<%=glosacampos%>,'MejorRespuesta','MejorFecha','MejorGesti�n');
		    identificadorfilas[tabla]="fila";
		    pievisible[tabla]=true;
		    columnavisible[tabla] = new Array(false,<% if checkvisible = "checked" then %> true <% else%> false <% end if%>,true<%=glosavisible%>,true,true,true);
		    anchocolumna[tabla] =  new Array('','2%','5%'<%=glosaancho%>,'5%','5%','10%');
		    aligncabecera[tabla] = new Array('left','left','left'<%=glosaaligncabecera%>,'left','left','left');
		    aligndetalle[tabla] = new Array('left','left','left'<%=glosaaligndetalle%>,'left','left','left');
		    alignpie[tabla] =     new Array('left','left','left'<%=glosaalignpie%>,'left','left','left');
		    decimalesnumero[tabla] = new Array(-1,-1,-1<%=glosadecimalesnumero%>,-1,-1,-1);
		    formatofecha[tabla] =   new Array('','',''<%=glosaformatofecha%>,'','','');


		    //Se escriben condiciones de datos administrados "objetos formulario"
		    idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
		    objetofomulario[tabla] = new Array();
		   		
				objetofomulario[tabla][0]='<input type=hidden name=idcampanapersona-id- value=-c0->' + '<a href="javascript:modificar(-id-);">-valor-</a>';
				objetofomulario[tabla][1]='<input type=checkbox name=-id- id=-id- value=-c0- onclick="javascript:checkarpersonas(-id-);">' + '-valor-';
				objetofomulario[tabla][2]='<a href="javascript:modificar(-id-);">-valor-</a>';				
				<%
				indicecampo=2
				Do while not RS.EOF 
				    indicecampo=indicecampo + 1
				    %>objetofomulario[tabla][<%=indicecampo%>]='<a href="javascript:modificar(-id-);">-valor-</a>';
				    <%
				RS.MoveNext 
				Loop
				RS.MoveFirst				
				%>
				objetofomulario[tabla][<%=(indicecampo +1)%>]='<a href="javascript:modificar(-id-);">-valor-</a>';
				objetofomulario[tabla][<%=(indicecampo +2)%>]='<a href="javascript:modificar(-id-);">-valor-</a>';
				objetofomulario[tabla][<%=(indicecampo +3)%>]='<a href="javascript:modificar(-id-);">-valor-</a>';						
					
		    filtrardatos[tabla]=0; //define si carga auto el filtro
		    filtrofomulario[tabla] = new Array();
		    tipofiltrofomulario[tabla] = new Array();
		    	filtrofomulario[tabla][0]='';	
		    	filtrofomulario[tabla][1]='';	  
		    	filtrofomulario[tabla][2]='';  
                <%
				indicecampo=2
				Do while not RS.EOF 
				    indicecampo=indicecampo + 1
				    %>filtrofomulario[tabla][<%=indicecampo%>]='';
				    <%
				RS.MoveNext 
				Loop
				RS.MoveFirst				
				%>
				filtrofomulario[tabla][<%=(indicecampo+1)%>]='';	  
		    	filtrofomulario[tabla][<%=(indicecampo+2)%>]='';  
		    	filtrofomulario[tabla][<%=(indicecampo+3)%>]='';  
				
				
		    valorfiltrofomulario[tabla] = new Array();
				valorfiltrofomulario[tabla][0]='';	
				valorfiltrofomulario[tabla][1]='';			
				valorfiltrofomulario[tabla][2]='';		
                <%
				indicecampo=2
				Do while not RS.EOF 
				    indicecampo=indicecampo + 1
				    %>valorfiltrofomulario[tabla][<%=indicecampo%>]='';
				    <%
				RS.MoveNext 
				Loop
				RS.MoveFirst				
				%>
				valorfiltrofomulario[tabla][<%=(indicecampo+1)%>]='';			
				valorfiltrofomulario[tabla][<%=(indicecampo+2)%>]='';	
				valorfiltrofomulario[tabla][<%=(indicecampo+3)%>]='';	

		    //Se escribe el conjunto de datos de tabla 0
		    datos[tabla]=new Array();
		<%
		
		'if buscador<>"" then
		''			filtrobuscador = "where a.IDCampa�a =" & idcampana & " and a.IDCampa�aPersona in ( select b.IDCampa�aPersona from Campa�a_Detalle a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where b.IDCampa�a = " & idcampana & " and IDCampa�aCampo = 1 and ValorTexto like '%" & buscador & "%')"
		''		end if'
		
		if filtrobuscador<>"" then
			filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
		end if		
				
		
		contadortotal=0
		
		'sql="select Count(*) / (select count(*) from Campa�a_Campo b inner join campa�a c on b.IDTipoCampa�a = 'c.IDTipoCampa�a where c.IDCampa�a =" & idcampana & " and b.FlagNroDocumento <> 1) from Campa�a_Detalle a where 'IDCampa�aPersona in (Select IDCampa�aPersona from Campa�a_Persona a " & filtrobuscador & " ) "  
		
		
		if trim(filtrobuscador2) <> "" then
		sql = "Select count(distinct IDCampa�aPersona) from Campa�a_Persona a " & filtrobuscador & " and a.IDCampa�aPersona IN (" & filtrobuscador2 & " )"
		else
		sql = "Select count(distinct IDCampa�aPersona) from Campa�a_Persona a " & filtrobuscador 
		end if
		'Response.Write filtrobuscador2
		'' response.write sql
		consultar sql,RS3
		contadortotal=rs3.fields(0)
		
		RS3.Close
		
		if contadortotal <> 0 then

				if trim(filtrobuscador2) <> "" then
				sql = "select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),  a.IDCampa�aPersona) from Campa�a_Persona a " & filtrobuscador & " and a.IDCampa�aPersona IN (" & filtrobuscador2 & " ) group by a.IDCampa�aPersona FOR XML PATH('') ), 1, 1, '') as Cadena"
				else
				sql = "select STUFF((SELECT CAST(',' AS varchar(MAX)) + CONVERT(VARCHAR(MAX),  a.IDCampa�aPersona) from Campa�a_Persona a " & filtrobuscador & " group by a.IDCampa�aPersona FOR XML PATH('') ), 1, 1, '') as Cadena"		
				end if


		'response.Write sql

				consultar sql,RS3

				idpersonas_asig = RS3.fields(0)

				RS3.Close

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

             
		
		       if ordentipo <> "" then
		       		
					sql="SELECT A.IDCampa�aPersona,A.NroDocumento, isnull((SELECT Usuario from Usuario where CodUsuario = A.UsuarioAsignado),'Sin asignar') as Asignacion from Campa�a_Persona A " & filtrobuscador & " order by A.IDCampa�aPersona" 
					
					if filtrobuscador2 <> "" or (filtrobuscador2 <>"" and buscador <>"") then
						
						sql="SELECT A.IDCampa�aPersona,A.NroDocumento, isnull((SELECT Usuario from Usuario where CodUsuario = A.UsuarioAsignado),'Sin asignar') as Asignacion from Campa�a_Persona A " & filtrobuscador & " and A.IDCampa�aPersona in (" & filtrobuscador2 & ") order by A.IDCampa�aPersona"					
					end if
		       else
					if pag>1 then					
					sql="SELECT TOP " & cantidadxpagina & " A.IDCampa�aPersona,A.NroDocumento , isnull((SELECT Usuario from Usuario where CodUsuario = A.UsuarioAsignado),'Sin asignar') as Asignacion from Campa�a_Persona A where A.idcampa�a=" & IDCampana & " and " & filtrobuscador1 & " A.IDCampa�aPersona NOT  IN (SELECT TOP " & topnovisible & " A.IDCampa�aPersona FROM Campa�a_Persona A  " & filtrobuscador & " order by A.IDCampa�aPersona) order by A.IDCampa�aPersona"		
					else
					sql="SELECT TOP " & cantidadxpagina & " A.IDCampa�aPersona,A.NroDocumento , isnull((SELECT Usuario from Usuario where CodUsuario = A.UsuarioAsignado),'Sin asignar') as Asignacion from Campa�a_Persona A " & filtrobuscador & " order by A.IDCampa�aPersona" 
					end if

					
					if filtrobuscador2 <> "" or (filtrobuscador2 <>"" and buscador <>"") then
						if pag>1 then			

						sql="SELECT TOP " & cantidadxpagina & " A.IDCampa�aPersona,A.NroDocumento ,isnull((SELECT Usuario from Usuario where CodUsuario = A.UsuarioAsignado),'Sin asignar') as Asignacion  from Campa�a_Persona A where " & filtrobuscador1 & " A.IDCampa�aPersona NOT  IN (SELECT TOP " & topnovisible & " A.IDCampa�aPersona FROM " & chr(10) & _
							"Campa�a_Persona A  " & filtrobuscador & " and A.IDCampa�aPersona in (" & filtrobuscador2 & ") ) and A.IDCampa�aPersona in (" & filtrobuscador2 & ")" & chr(10) & _
							" order by A.IDCampa�aPersona"
						else
						sql="SELECT TOP " & cantidadxpagina & " A.IDCampa�aPersona,A.NroDocumento , isnull((SELECT Usuario from Usuario where CodUsuario = A.UsuarioAsignado),'Sin asignar') as Asignacion from Campa�a_Persona A " & filtrobuscador & " and A.IDCampa�aPersona in (" & filtrobuscador2 & ") order by A.IDCampa�aPersona"	
						end if
					end if
				end if
		''response.write sql

				consultar sql,RS3
				contador=0
		 
		
		if RS.RecordCount > 0 then
					idpersonas =""
					Do while not RS3.EOF 
						if idpersonas = "" then
							idpersonas = RS3.Fields("IDCampa�aPersona")
						Else
							idpersonas = idpersonas & "," & RS3.Fields("IDCampa�aPersona")
						end if				
					RS3.MoveNext 
					Loop 			
		

				if ordentipo <> "" then
					RS3.Close
					
						             if pag>1 then
						             		  sql = "select TOP " & cantidadxpagina & " a.IDCampa�aPersona, b.NroDocumento, isnull((SELECT Usuario from Usuario where CodUsuario = b.UsuarioAsignado),'Sin asignar') as Asignacion from campa�a_detalle  a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where a.IDCampa�aCampo = " & ordencampo & " and a.IDCampa�aPersona in (" & idpersonas & ") and a.IDCampa�aPersona NOT  IN (SELECT TOP " & topnovisible & " a.IDCampa�aPersona from campa�a_detalle  a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where a.IDCampa�aCampo = " & ordencampo & " and a.IDCampa�aPersona in (" & idpersonas & ")  order by a.ValorTexto " & ordentipo  & ",a.ValorEntero " & ordentipo  & ",a.ValorFloat " & ordentipo  & ",a.valorfecha " & ordentipo & " ) order by a.ValorTexto " & ordentipo  & ",a.ValorEntero " & ordentipo  & ",a.ValorFloat " & ordentipo  & ",a.valorfecha " & ordentipo  
						             else
						               sql = "select TOP " & cantidadxpagina & " a.IDCampa�aPersona, b.NroDocumento, isnull((SELECT Usuario from Usuario where CodUsuario = b.UsuarioAsignado),'Sin asignar') as Asignacion from campa�a_detalle  a inner join Campa�a_Persona b on a.IDCampa�aPersona = b.IDCampa�aPersona where a.IDCampa�aCampo = " & ordencampo & " and a.IDCampa�aPersona in (" & idpersonas & ")  order by a.ValorTexto " & ordentipo  & ",a.ValorEntero " & ordentipo  & ",a.ValorFloat " & ordentipo  & ",a.valorfecha " & ordentipo  
						             end if

						  '' response.Write  "querty:" & sql

					consultar sql, RS3
					

					if RS3.RecordCount > 0 then
					idpersonas =""
					Do while not RS3.EOF 
						if idpersonas = "" then
							idpersonas = RS3.Fields("IDCampa�aPersona")
						Else
							idpersonas = idpersonas & "," & RS3.Fields("IDCampa�aPersona")
						end if				
					RS3.MoveNext 
					Loop 			
					end if
					RS3.MoveFirst
				else
					RS3.MoveFirst
				end if




		   		sql="select A.IDCampa�aPersona,A.NroDocumento,C.IDCampa�aCampo,C.NroCampo,C.TipoCampo,D.ValorTexto,D.ValorEntero,D.ValorFloat,D.ValorFecha " & chr(10) & _
		            "from Campa�a_Persona A " & chr(10) & _
		            "inner join Campa�a B " & chr(10) & _
		            "on A.IDCampa�a=B.IDCampa�a " & chr(10) & _
		            "inner join Campa�a_Campo C " & chr(10) & _
		            "on B.IDTipoCampa�a=C.IDTipoCampa�a and C.Nivel=1 and C.Visible=1 " & chr(10) & _
		            "inner join Campa�a_Detalle D " & chr(10) & _
		            "on A.IDCampa�aPersona=D.IDCampa�aPersona and C.IDCampa�aCampo=D.IDCampa�aCampo " & chr(10) & _
		            " where A.IDCampa�aPersona IN (" & idpersonas & ")"
		           '' response.write sql
				consultar sql,RS2

		
				Do while not RS3.EOF		
									
					%>
					datos[tabla][<%=contador%>] = new Array();
						datos[tabla][<%=contador%>][0]=<%=RS3.Fields("IDCampa�aPersona")%>;	
						datos[tabla][<%=contador%>][1]='';		
						datos[tabla][<%=contador%>][2]='<%=RS3.Fields("Asignacion")%>';			
		                <%
						indicecampo=2
						Do while not RS.EOF
						    indicecampo=indicecampo + 1
		   				    if RS.Fields("FlagNroDocumento")=0 then
						        RS2.Filter=" IDCampa�aPersona=" & RS3.Fields("IDCampa�aPersona") & " and IDCampa�aCampo=" & RS.Fields("IDCampa�aCampo") & " "
		    				    
						        Select Case RS.Fields("TipoCampo")
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
						        Response.Write "datos[tabla][" & contador & "][" & indicecampo & "]='" & RS3.Fields("NroDocumento") & "';" & chr(10)
						    end if						   	
						RS.MoveNext 
						Loop
						RS.MoveFirst
						%>
						 datos[tabla][<%=contador%>][<%=(indicecampo+1)%>]='<%
						 sql = "select top 1 c.Descripcion from Campa�a_Persona_Accion b inner join Gestion c on b.IDGestion = c.IDGestion where IDCampa�aPersona = " & RS3.Fields("IDCampa�aPersona") & " order by prioridad asc , b.FechaRegistra desc"

						 consultar sql, RS4

						 if RS4.RecordCount > 0 then

						 Response.write RS4.fields("Descripcion")

						else

						response.write ""
						end if

						 RS4.Close

						 %>';	
						 datos[tabla][<%=contador%>][<%=(indicecampo+2)%>]='<%

						 sql = "select top 1 b.FechaRegistra from Campa�a_Persona_Accion b " & chr(10) & _
							"inner join Gestion c on b.IDGestion = c.IDGestion" & chr(10) & _
							"where IDCampa�aPersona = " & RS3.Fields("IDCampa�aPersona") & chr(10) & _
							"order by prioridad asc , b.FechaRegistra desc"

						 consultar sql, RS4

						  if RS4.RecordCount > 0 then

						 Response.write RS4.fields("FechaRegistra")

						else

						response.write ""
						end if

						 RS4.Close

						 %>';		
						 datos[tabla][<%=contador%>][<%=(indicecampo+3)%>]='<%

						 sql = "select top 1 b.Comentario from Campa�a_Persona_Accion b " & chr(10) & _
							"inner join Gestion c on b.IDGestion = c.IDGestion " & chr(10) & _
							"where IDCampa�aPersona = " & RS3.Fields("IDCampa�aPersona") & chr(10) & _
							"order by prioridad asc , b.FechaRegistra desc"

							 consultar sql, RS4

						  if RS4.RecordCount > 0 then

						 Response.write RS4.fields("Comentario")

						else

						response.write ""
						end if

						 RS4.Close
						 %>';					
						 <%
				contador=contador + 1
				RS3.MoveNext 
				Loop 			
				RS3.Close

			end if
		end if 
		%>
			    
				    //datos del pie si fuera visible
				    pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;'<%=glosapie%>,'&nbsp;','&nbsp;','&nbsp;');
				    piefunciones[tabla] = new Array('','',''<%=glosapiefunciones%>,'','',''); 


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
				
		sql= "select Descripcion from campa�a where IDCampa�a = " & IDCampana
		consultar sql,RS3
		descripcioncampana = rs3.fields("Descripcion")
		RS3.Close

	


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
							''sql = "select ROW_NUMBER() OVER(ORDER BY IDCampa�aCampo ) AS nro ,IDCampa�aCampo, GlosaCampo, TipoCampo, FlagNroDocumento from Campa�a_Campo a inner join Campa�a b on a.IDTipoCampa�a = b.IDTipoCampa�a where a.Nivel = 1 and b.IDCampa�a = " & IDCampana
							''consultar sql,RS
							
							cadenareset = ""
							varcolor = 1
							Do While Not RS.EOF								
						%>

						<tr class="fondo-red <% IF varcolor <> 0 Then %> fondo-blanco <% Else %> fondo-rojo <% End IF %>" >
							<input type="hidden" name="idcampanacampo<%=RS.fields("IDCampa�aCampo")%>" value="<%=obtener("idcampanacampo" & RS.fields("IDCampa�aCampo"))%>">							
							<input type="hidden" name="idTipocampo<%=RS.Fields("IDCampa�aCampo")%>" value="<%If RS.Fields("FlagNroDocumento") = 0 then%><%=RS.Fields("TipoCampo")%><%Else%>0<%End if%>">
							<td>							
							<%=RS.Fields("GlosaCampo")%>
							</td>
							<td class="text-withe" width="120" style="text-align: center;">
								<%cadenareset = cadenareset & "document.formula.filtrado" & RS.fields("IDCampa�aCampo") & ".value=0;" & chr(13)%>
									<select name="filtrado<%=RS.fields("IDCampa�aCampo")%>" id="Select<%=RS.Fields("IDCampa�aCampo")%>" onChange="javascript:modfiltro++;vercajatexto2('<%=trim(RS.fields("IDCampa�aCampo")) & "b"%>','Select<%=RS.Fields("IDCampa�aCampo")%>');" style="font-size: xx-small; text-align: center; width: 100px;">
										<option value="0">Seleccione un filtro</option>
										<%
										sql = "SELECT idfiltro, descripcion FROM Filtro WHERE TipoCampo =" & RS.Fields("TipoCampo") 
										consultar sql,RS4
										Do While Not  RS4.EOF
										%>	
											<option value="<%=RS4.Fields("idfiltro")%>" 
												<% if obtener("filtrado" & RS.fields("IDCampa�aCampo")) <> "" Then
												 if RS4.fields("idfiltro") = CInt(obtener("filtrado" & RS.fields("IDCampa�aCampo"))) Then %> 
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
								<%cadenareset = cadenareset & "document.formula.buscador2" & RS.fields("IDCampa�aCampo") & ".value='';" & chr(13)%>
								<input type="text"  onChange="javascript:modfiltro++;" name="buscador2<%=RS.Fields("IDCampa�aCampo")%>"  class="form-control" value="<%=obtener("buscador2" & RS.Fields("IDCampa�aCampo"))%>"/>
								<% 
								if RS.Fields("TipoCampo") = "2" or RS.Fields("TipoCampo") ="3" or RS.Fields("TipoCampo") = "4" then 									
								%>

								<%cadenareset = cadenareset & "document.formula.buscador2" & RS.fields("IDCampa�aCampo") & "b.value='';" & chr(13)%>
								<input type="text"  onChange="javascript:modfiltro++;" id="<%=RS.Fields("IDCampa�aCampo")%>b"
								<% if obtener("filtrado" & RS.fields("IDCampa�aCampo")) <> "" Then 
								if CInt(obtener("filtrado" & RS.fields("IDCampa�aCampo"))) <> 12 and CInt(obtener("filtrado" & RS.fields("IDCampa�aCampo"))) <> 8 and CInt(obtener("filtrado" & RS.fields("IDCampa�aCampo"))) <> 14 Then 
								%>
								 style="display:none;" <% end if 
								 else %>
								  style="display:none;"
								 <%end if %>name="buscador2<%=RS.Fields("IDCampa�aCampo")%>b"  class="form-control" value="<%=obtener("buscador2" & RS.Fields("IDCampa�aCampo") & "b")%>"/>
								 <% end if %>
							</td>													
						</tr>
						<%
							if varcolor <> 0 then
								varcolor = 0
							else
							 	varcolor = 1
						    end if
							RS.MoveNext
							loop
							RS.MoveFirst
							%>
							<tr class="fondo-red">
								<td class="text-withe" width="120" style="background: #FE6D2E;" colspan="4" >	
								Filtro por Gesti�n.
							</td>
							</tr>
							<tr class="fondo-red">
								<td class="text-withe" width="120" style="background: #FE6D2E;" colspan="2">	
								Respuesta de Gesti�n:
							</td>
							<td style="background: #FE6D2E;" colspan="2">
								<select name="codrespuesta" id="codrespuesta" onChange="javascript:modfiltro++;">
									<option value="">Debe escoger una Respuesta.</option>
								<%
								sql = "select a.IDGestion, a.Descripcion from Gestion a inner join Campa�a b on b.IDTipoCampa�a = a.IDTipoCampa�a where b.IDCampa�a = " & idcampana

								consultar sql, RS4

								DO While Not RS4.EOF
								%>
								<option value="<%=RS4.Fields("IDGestion")%>" <% if Cstr(RS4.Fields("IDGestion")) = codrespuesta Then  %> selected <% else %> <% end if %>><%=RS4.Fields("Descripcion")%></option>
								<%
								RS4.MoveNext
								loop
								RS4.Close
								%>
								</select>
							</td>
							</tr>
					</table>			
					<script type="text/javascript">
						function limpiarform()
						{
							<%=cadenareset%>
							document.formula.codrespuesta.value ="";
							modfiltro++;
						}
					</script>			
				</div>		

				<div id="modal-filtro2" class="filtro-visible2 no-visible" >	
					<table border="0">
						<tr class="fondo-red">
							<td class="text-withe">
								Asignar Filtro.
							</td>
							<td id="close-modal2" style="text-align: right;" ><a style="float:right; padding-right:0px;" href="#"><i style="color: white;" class="demo-icon2 icon-cancel-circle">&#xe807;</i></a></td>
							
						</tr>
						<tr class="fondo-red  fondo-blanco ">
							<td><font size="2">Campa�a:</font></td>
							<td><font size="2"><%=Nombcampana%><br>Inicio:<%=fechainicio%>&nbsp;al:<%=fechafin%></font></td>
						</tr>
						<tr class="fondo-red  fondo-blanco ">
						<td><font size="2">Asignar a:</font></td>
						<td>
							<select name="codusuario" style="font-size: xx-small; width: 200px;">
							<OPTION value="">Seleccione un Operador</OPTION>
							<%
							sql = "SELECT CodUsuario, Usuario FROM Usuario"
							consultar sql,RS4
							Do While Not  RS4.EOF
							%>
								<option value="<%=RS4.Fields("CodUsuario")%>" <% if codusuario<>"" then%><% if RS4.fields("CodUsuario")=int(codusuario) then%> selected<%end if%><%end if%>><%=RS4.Fields("Usuario")%></option>
							<%
							RS4.MoveNext
							loop
							RS4.Close
							%>
							</select>
						</td>
					</tr>	
					<tr class="fondo-red" style="text-align: right;">
						<td colspan="2">
							<%
								if trim(filtrobuscador2) <> "" then
								sql = "select  count(a.IDCampa�aPersona) AS Nro from Campa�a_Persona a " & filtrobuscador & " and a.UsuarioAsignado is not null and a.IDCampa�aPersona IN (" & filtrobuscador2 & " ) "
							    else
								sql = "select Count(a.IDCampa�aPersona) as Nro from Campa�a_Persona a " & filtrobuscador & " and a.UsuarioAsignado is not null "
								end if
								
								consultar sql,RS4

								cruces= RS4.Fields("Nro")
								RS4.Close
							%>
							<% if checkvisible <> "checked" then %>
								<a href="javascript:asignar('<%=cruces%>');"><i class="demo-icon icon-floppy">&#xe809;</i></a>
							<% else %>
							    <a href="javascript:asignar(0);"><i class="demo-icon icon-floppy">&#xe809;</i></a>
							<% end if %>
								&nbsp;
							</td>
					</tr>
					</table>
					
				</div>
				<table width="100%" cellpadding="4" cellspacing="0" border="0"><!--Esto no sale -->	
					<tr class="fondo-orange">
						<td class="text-orange" align="left" width="250"><font size="2" face="Fira Sans Condensed"><b><%if contadortotal=0 then%><%=descripcioncampana%> (0) - No hay registros.<%Else%><%=descripcioncampana%> (<%=contadortotal%>)<%end if%></b></font></td>
							<td class="text-orange" align="right" width="250"><font size="2" face="Fira Sans Condensed">Buscar:&nbsp;<input name="textobuscar" value="<%=buscador%>" size="20" id="textobuscar" onkeypress="if(window.event.keyCode==13) buscar();"></font></td>
							<td width="80">
								<a href="javascript:buscar();"><i class="demo-icon icon-search">&#xe80c;</i></a>
								<a id="show-filtro" href="#"><i class="demo-icon icon-filter">&#xe820;</i></a>							
							</td>
							<td width="100">Ver:
								<SELECT name="paginado" onChange="javascript:buscar();">
								<option value="18" <%if CInt(paginado) = 18 then %>selected<%end if%> >18</option>
								<option value="50" <%if CInt(paginado) = 50 then %>selected<%end if%> >50</option>
								<option value="100" <%if CInt(paginado) = 100 then %>selected<%end if%>  >100</option>
								<option value="500" <%if CInt(paginado) = 500 then %>selected<%end if%> >500</option>
								<option value="1000" <%if CInt(paginado) = 1000 then %>selected<%end if%> >1000</option>
								<!-- <option value="<%=contadortotal%>" <%if CInt(paginado) = contadortotal then %>selected<%end if%> >Todos</option> -->
								</select>					
							</td>
							<td width="170"> Ordenar[
								<Select name="ordencampo" style="width: 50px;"  id="ordencampo" onChange="javascript:habilitarorden();" >
									<option value="">Seleccione Campo Orden</option>
								<%	Do While Not RS.EOF								
								%>
									<option value="<%=RS.Fields("IDCampa�aCampo")%>" <%if ordencampo =  RS.Fields("IDCampa�aCampo") then %>selected<%end if%>><%=RS.Fields("GlosaCampo")%></option>
								<%
									RS.MoveNext
									loop
									RS.MoveFirst
								%>
								</Select>
								<Select  style="width: 50px; float: absolute; " name="ordentipo" id="ordentipo" onChange="javascript:buscar();" <%if ordencampo=0 then%> disabled <%end if%>>
									<option value="">Ord</option>
									<option value="asc" <%if ordentipo = "asc" then %> selected <%end if%>>Ascendente</option>
									<option value="desc" <%if ordentipo = "desc" then %> selected <%end if%>>Descendente</option>
								</Select>]
								
							</td>																									
							<td width="120">
								Asignado[
								<Select style="width: 50px; " name="filoperador"  id="filoperador" onChange="javascript:buscar();">
								<option value="">Todos</option>
								<% sql=" select distinct isNull(UsuarioAsignado,0) as codusuario, ISNULL((select Usuario from Usuario where CodUsuario = UsuarioAsignado),'Sin Asignar') as Usuario from Campa�a_Persona where IDCampa�a = " & IDCampana

								consultar sql,RS4
								do while not RS4.EOF
								%>
								<option value="<%=RS4.fields("codusuario")%>" 	<% if filoperador <>"" then 
								IF Cint(filoperador) = Cint(RS4.fields("codusuario")) then %> Selected <%
								END IF 
								end if%>><%=RS4.fields("Usuario")%>							 
								</option>
								<%
									RS4.MoveNext
									loop
									RS4.Close
								%>
								</Select>]
							</td>
							<td style="vertical-align: middle;">
								<span style="margin-left: -4px; float: left; margin-top: 2.2px;">Sel[</span>
								
								<input type="checkbox" name="checkvisible" style="margin-left: 1px; float: left; vertical-align: middle;"  id="checkvisible" <%=checkvisible%> value="<%=checkvisible%>" onclick="javascript:activarchecks(this);">
								<input style="vertical-align: middle; float: left; margin-left: -1px;" type="checkbox" name="seltodo" id="seltodo" onclick="javascript:marcar(this)">
								<span style="margin-left: -4px; float: left; margin-top: 2.2px;">]</span>							
							</td>	
						<td class="text-orange" align="right">
							&nbsp;&nbsp;<a  id="show-filtro2" href="#"><i class="demo-icon icon-cog">&#xe81f;</i></a>
							&nbsp;&nbsp;<a href="javascript:exportar();"><i class="demo-icon icon-file-excel">&#xf1c3;</i></a>
						<%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><i class="demo-icon icon-download">&#xe814;</i></a><%end if%></b></font>
						</td>
						<%if contadortotal>0 then%>
								<td class="text-orange" align="right" width="180"><font size="2" face="Fira Sans Condensed">P�g.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
						<%end if%>
					</tr>	
				</table>
				<div id="tabla0"> 
				</div>
		
		
		<input type="hidden" name="actualizarlista" value="">
		<input type="hidden" name="expimp" value="">	
		<input type="hidden" name="idpersonas_asig" value="<%=idpersonas_asig%>">
		<input type="hidden" name="idpersonas_chk" id="idpersonas_chk" value="">
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
		alert("Ud. No tiene autorizaci�n para este proceso.");
		window.open("dcs_userexpira.asp","_top");
	</script>
	<%	
	
	end if	
	
	if contador > 0 then
	RS2.Close
	end if
	RS.Close
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



