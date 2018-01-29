  //variable generales
    var nombreformulario="formula";
    var ascendentetmp="";
    var ordentmp=0;
    var tabla;
    var comilladoble='"';
    var comillasimple="'";
    var datoscargados=0;
    var rutaimgcab="";


  //Variables de arreglos de tabla
    var orden=new Array();
    var ascendente=new Array();
    var nrocolumnas=new Array();
    var fondovariable=new Array();
    var anchotabla=new Array();
    var botonfiltro=new Array();
    var botonactualizar=new Array();
    var botonagregar=new Array();
    var spacingtabla=new Array();
    var paddingtabla=new Array();
    var cabecera = new Array();
    var seteocabecera = new Array();
    var identificadorseteo = new Array();
    var identificadorfilas=new Array();
    var pievisible = new Array();
    var columnavisible = new Array();
    var anchocolumna = new Array();
    var aligncabecera = new Array();
    var aligndetalle = new Array();
    var alignpie = new Array();
    var piefunciones = new Array();
    var decimalesnumero = new Array();
    var formatofecha = new Array();
    var idobjetofomulario=new Array();
    var objetofomulario = new Array(); 
    var nombreobjetofomulario = new Array(); 
    var filtrardatos = new Array(); 
    var filtrofomulario = new Array(); 
    var valorfiltrofomulario = new Array(); 
    var tipofiltrofomulario = new Array(); 
    var datos = new Array(); 
    var pievalores = new Array();
    var posicionselect = new Array();
    var nombreselect = new Array();
    var opcionesvalor = new Array();
    var opcionestexto = new Array();
    var funcionactualiza = new Array();
    var funcionagrega = new Array();
    var usofiltrocheckbox= new Array(); //SE USA INTERNO: CUANDO HAY MAS DE UN CHECK BOX EN EL FILTRO

    //Funciones por conjunto de datos
    function dibujarTabla(x) 
    {
	datoscargados=0;
        var fondofila;

        var html="";
	
	//boton para filtros
	if(botonfiltro[x])
	{
		if (filtrardatos[x]==1)
		{
			html+="<input type=button value='Quitar Filtros' style='font-size: xx-small;' onclick='filtrardatos[" + x + "]=0;dibujarTabla(" + x + ");'>";
		}
		else
		{
			html+="<input type=button value='Filtrar Datos' style='font-size: xx-small;' onclick='filtrardatos[" + x + "]=1;dibujarTabla(" + x + ");'>";
			borrarfiltros(x);
		}
	}
	if(botonactualizar[x])
	{
		html+="<input type=button value='Actualizar Datos' style='font-size: xx-small;' onclick='" + funcionactualiza[x] + "'>";
	}
	if(botonagregar[x])
	{
		html+="<input type=button value='Agregar Datos' style='font-size: xx-small;' onclick='" + funcionagrega[x] + "'>";
	}
	html += "<table width=" + anchotabla[x] + " cellpadding='" + paddingtabla[x] + "' cellspacing='" + spacingtabla[x] + "'>";
            
	html += '<tr>';
	for (var i=0; i<nrocolumnas[x]; i++) 
	{
	    if (columnavisible[x][i])
	    {
		var seteo='';
		if(seteocabecera.length>x)
		{
		    switch(seteocabecera[x][i])
		    {
			case "checkbox": 
				seteo='<input type=checkbox name="seteo_' + x.toString() + '_' + i.toString() + '" style="font-size: xx-small;" onclick="setearvisibles(' + x + ',' + i + ',this);window.event.cancelBubble=true;">';
				break;
			case "select": 
				seteo='<select name="seteo_' + x.toString() + '_' + i.toString() + '" style="font-size: xx-small;" onchange="setearvisibles(' + x + ',' + i + ',this);" onclick="window.event.cancelBubble = true;"></select>';
				break;
			default: seteo='';
		    }
		}
	        html += '<th onclick="cambiar(' + i + ',' + x + ')" align="' + aligncabecera[x][i] + '" width="' + anchocolumna[x][i] + '">';
        	html += ((orden[x] == i)?(ascendente[x]?  '<i class="demo-icon3 icon-up-open">&#xe817;</i>': '<i class="demo-icon3 icon-down-open">&#xe816;</i>'): '<i class="demo-icon3 icon-sort">&#xf0dc;</i>');
	        html += cabecera[x][i] + seteo + '</th>';
	    }
	}
        html += '</tr>';

	//filtros
	if(filtrardatos[x]==1)
	{
		html += '<tr>';
		for (var i=0; i<nrocolumnas[x]; i++) 
		{
		    if (columnavisible[x][i])
		    {
	        	html += '<td align="' + aligncabecera[x][i] + '" width="' + anchocolumna[x][i] + '">' + reemplazartexto((reemplazartexto(filtrofomulario[x][i],'-valorfiltro-',valorfiltrofomulario[x][i])),'-columna-',i) + '</td>';
		    }
		}
	        html += '</tr>';
	}

	//Ordena arreglo
	ascendentetmp=ascendente[x];
	ordentmp=orden[x];
	datos[x].sort(organizar);
    
	var contador=0;
	var mostrarfila=0;
	var htmlfila=0;
        for (var i=0; i<datos[x].length; i++) 
	{
            htmlfila="";
	    htmlfila += '<tr';
	    if (identificadorfilas[x]!="") htmlfila += ' id="' + identificadorfilas[x] + datos[x][i][idobjetofomulario[x]] + '"';
	    htmlfila += '>'

 	    mostrarfila=1;
	    if(contador%2==0) fondofila='';
	    else fondofila=fondovariable[x];

       	    for (var j=0; j<nrocolumnas[x]; j++) 
	    {
	      //seteamos los valores del pie si hay funcion
	      if(i==0 && (piefunciones[x][j]=="suma" || piefunciones[x][j]=="promedio")) pievalores[x][j]=0;

	      if (columnavisible[x][j])
	      {
		   var valor;
		   var id;
	
		   if(formatofecha[x][j]!="")
		   {
		      if(datos[x][i][j]!=null)
		      {
			valor=formatofecha[x][j];
			if(datos[x][i][j].getDate()>9) valor=reemplazartexto(valor,"dd",datos[x][i][j].getDate());
			else valor=reemplazartexto(valor,"dd","0" + datos[x][i][j].getDate());
			if(datos[x][i][j].getMonth() + 1>9) valor=reemplazartexto(valor,"mm",datos[x][i][j].getMonth() + 1);
			else valor=reemplazartexto(valor,"mm","0" + (datos[x][i][j].getMonth() + 1));
			valor=reemplazartexto(valor,"aaaa",datos[x][i][j].getFullYear());
			if(datos[x][i][j].getHours()>9) valor=reemplazartexto(valor,"HH",datos[x][i][j].getHours());
			else valor=reemplazartexto(valor,"HH","0" + datos[x][i][j].getHours());
			if(datos[x][i][j].getMinutes()>9) valor=reemplazartexto(valor,"MI",datos[x][i][j].getMinutes());
			else valor=reemplazartexto(valor,"MI","0" + datos[x][i][j].getMinutes());
			if(datos[x][i][j].getSeconds()>9) valor=reemplazartexto(valor,"SS",datos[x][i][j].getSeconds());
			else valor=reemplazartexto(valor,"SS","0" + datos[x][i][j].getSeconds());
			if (valor.indexOf(NaN)>0) valor="";
		      }
		      else valor="";
		   }
		   else
		   {
		      if(datos[x][i][j]!=null)
		      {
			   if(decimalesnumero[x][j]>=0) valor=FormatNumber(datos[x][i][j],decimalesnumero[x][j]);
			   else valor=datos[x][i][j];
		      }
		      else valor="";
		   }
		   //se agrega x filtro
	           if (filtrardatos[x]==1)
		   {
		        if(valorfiltrofomulario[x][j]!="")
			{
  		           //se agrega x filtro comparativo ejm. >=3 y < 5
   			   var filtroY=false;
   			   var filtroO=false;
   			   var mayorigualque;
			   var mayorque;
			   var menorigualque;
			   var menorque;
			   var diferenteque;
			   var diferentequetext="";
			   var escomparacion=false;
			   var filtrarpor=valorfiltrofomulario[x][j];

                           //rutina cargar mayorigualque, mayor que,...
			   if(filtrarpor.substring(0,2)=="<>" && filtrarpor.replace("<>","").length>0)
			   {
			      if(!isNaN(parseFloat(filtrarpor.substring(2,filtrarpor.length))))
			      {
				diferenteque=parseFloat(filtrarpor.substring(2,filtrarpor.length));
				escomparacion=true;
			      }	
			      else
			      {
				diferentequetext=filtrarpor.substring(2,filtrarpor.length);
				escomparacion=true;				
			      }		      
			   }
			   else
			   {
		              if((filtrarpor.substring(0,1)==">" || filtrarpor.substring(0,1)=="<") && filtrarpor.replace(">=","").replace("<=","").replace("<","").replace(">","").length>0)
			      {
				if(filtrarpor.substring(0,2)==">=")
				{
				   if(!isNaN(parseFloat(filtrarpor.substring(2,filtrarpor.length))))
				   {
					mayorigualque=parseFloat(filtrarpor.substring(2,filtrarpor.length));
					escomparacion=true;
				   }
				}
				else
				{
				   if(filtrarpor.substring(0,2)=="<=")
				   {
				      if(!isNaN(parseFloat(filtrarpor.substring(2,filtrarpor.length))))
				      {
					menorigualque=parseFloat(filtrarpor.substring(2,filtrarpor.length));
					escomparacion=true;
				      }
				   }
				   else
				   {
				      if(filtrarpor.substring(0,1)==">")
				      {	
				   	 if(!isNaN(parseFloat(filtrarpor.substring(1,filtrarpor.length))))
					 {
					   mayorque=parseFloat(filtrarpor.substring(1,filtrarpor.length));
					   escomparacion=true;
				         }
				      }
				      else
				      {
				      	 if(filtrarpor.substring(0,1)=="<")
				         {	
				   	   if(!isNaN(parseFloat(filtrarpor.substring(1,filtrarpor.length))))
					   {
					      menorque=parseFloat(filtrarpor.substring(1,filtrarpor.length));
					      escomparacion=true;
				           }
				         }					
				      }			
				   }
				}
			      }
   			   }

			   if(diferentequetext=="")
			   {
			     if(filtrarpor.toLowerCase().indexOf(" y ")>0)
			     {
				filtroY=true;
				filtrarpor=filtrarpor.substring(filtrarpor.toLowerCase().indexOf(" y ") + 3,filtrarpor.length);
			     }
			     else
			     {
			       if(filtrarpor.toLowerCase().indexOf(" o ")>0)
			       {
				 filtroO=true;
				 filtrarpor=filtrarpor.substring(filtrarpor.toLowerCase().indexOf(" o ") + 3,filtrarpor.length);
			       }
			     }

			     if(filtroY || filtroO)
			     {
				   //rutina cargar mayorigualque, mayor que,...
				   if(filtrarpor.substring(0,2)=="<>" && filtrarpor.replace("<>","").length>0)
				   {
				      if(!isNaN(parseFloat(filtrarpor.substring(2,filtrarpor.length))))
				      {
					diferenteque=parseFloat(filtrarpor.substring(2,filtrarpor.length));
					escomparacion=true;
				      }	
				      else
				      {
					diferentequetext=filtrarpor.substring(2,filtrarpor.length);
					escomparacion=true;				
				      }		      
				   }
				   else
				   {
			              if((filtrarpor.substring(0,1)==">" || filtrarpor.substring(0,1)=="<") && filtrarpor.replace(">=","").replace("<=","").replace("<","").replace(">","").length>0)
				      {
					if(filtrarpor.substring(0,2)==">=")
					{
					   if(!isNaN(parseFloat(filtrarpor.substring(2,filtrarpor.length))))
					   {
						mayorigualque=parseFloat(filtrarpor.substring(2,filtrarpor.length));
						escomparacion=true;
					   }
					}
					else
					{
					   if(filtrarpor.substring(0,2)=="<=")
					   {
					      if(!isNaN(parseFloat(filtrarpor.substring(2,filtrarpor.length))))
					      {
						menorigualque=parseFloat(filtrarpor.substring(2,filtrarpor.length));
						escomparacion=true;
					      }
					   }
					   else
					   {
					      if(filtrarpor.substring(0,1)==">")
					      {	
					   	 if(!isNaN(parseFloat(filtrarpor.substring(1,filtrarpor.length))))
						 {
						   mayorque=parseFloat(filtrarpor.substring(1,filtrarpor.length));
						   escomparacion=true;
					         }
					      }
					      else
					      {
					      	 if(filtrarpor.substring(0,1)=="<")
					         {	
				   		   if(!isNaN(parseFloat(filtrarpor.substring(1,filtrarpor.length))))
						   {
						      menorque=parseFloat(filtrarpor.substring(1,filtrarpor.length));
						      escomparacion=true;
					           }
					         }					
					      }			
					   }
					}
				      }
	   			   }				
			     }
			   }

			   if(!escomparacion)
			   {
				//este el el filtro directo si es igual o contiene
			   	if(tipofiltrofomulario[x][j]!="igual")  
				   {			   
				   	if(valor.toString().toLowerCase().indexOf(valorfiltrofomulario[x][j].toLowerCase())<0) mostrarfila=0;
				   }
				   else
				   {		
				   	if(valor.toString().toLowerCase()!=valorfiltrofomulario[x][j].toLowerCase()) mostrarfila=0;
				   }								
			   }
			   else
			   {
				/*
				Caso 1: <> Texto No Contiene
				Case 2 filtro y/o false: <>,>=,<=,>,<
				Case 3 filtro y/o true : <> y/o >=, etc.
				*/
				if(diferentequetext!="")
				{
				  if(valor.toString().toLowerCase().indexOf(diferentequetext.toLowerCase())>=0) mostrarfila=0;
				}
				else
				{
				  if(!filtroY && !filtroO)
				  {
				     if(!isNaN(mayorigualque) && parseFloat(valor)<mayorigualque) mostrarfila=0;
				     if(!isNaN(menorigualque) && parseFloat(valor)>menorigualque) mostrarfila=0;
				     if(!isNaN(mayorque) && parseFloat(valor)<=mayorque) mostrarfila=0;
				     if(!isNaN(menorque) && parseFloat(valor)>=menorque) mostrarfila=0;
				     if(!isNaN(diferenteque) && parseFloat(valor)==diferenteque) mostrarfila=0;
				  }
				  else
				  {
				     if(filtroY)	
				     {
				     	if((!isNaN(mayorigualque) && !isNaN(menorigualque)) && (parseFloat(valor)<mayorigualque || parseFloat(valor)>menorigualque)) mostrarfila=0;
					if((!isNaN(mayorigualque) && !isNaN(mayorque)) && (parseFloat(valor)<mayorigualque || parseFloat(valor)<=mayorque)) mostrarfila=0;
					if((!isNaN(mayorigualque) && !isNaN(menorque)) && (parseFloat(valor)<mayorigualque || parseFloat(valor)>=menorque)) mostrarfila=0;
					if((!isNaN(mayorigualque) && !isNaN(diferenteque)) && (parseFloat(valor)<mayorigualque || parseFloat(valor)==diferenteque)) mostrarfila=0;

					if((!isNaN(menorigualque) && !isNaN(mayorque)) && (parseFloat(valor)>menorigualque || parseFloat(valor)<=mayorque)) mostrarfila=0;
					if((!isNaN(menorigualque) && !isNaN(menorque)) && (parseFloat(valor)>menorigualque || parseFloat(valor)>=menorque)) mostrarfila=0;
					if((!isNaN(menorigualque) && !isNaN(diferenteque)) && (parseFloat(valor)>menorigualque || parseFloat(valor)==diferenteque)) mostrarfila=0;

					if((!isNaN(mayorque) && !isNaN(menorque)) && (parseFloat(valor)<=mayorque || parseFloat(valor)>=menorque)) mostrarfila=0;
					if((!isNaN(mayorque) && !isNaN(diferenteque)) && (parseFloat(valor)<=mayorque || parseFloat(valor)==diferenteque)) mostrarfila=0;

					if((!isNaN(menorque) && !isNaN(diferenteque)) && (parseFloat(valor)>=menorque || parseFloat(valor)==diferenteque)) mostrarfila=0;
				     }
				     else
				     {
				     	if((!isNaN(mayorigualque) && !isNaN(menorigualque)) && (parseFloat(valor)<mayorigualque && parseFloat(valor)>menorigualque)) mostrarfila=0;
					if((!isNaN(mayorigualque) && !isNaN(mayorque)) && (parseFloat(valor)<mayorigualque && parseFloat(valor)<=mayorque)) mostrarfila=0;
					if((!isNaN(mayorigualque) && !isNaN(menorque)) && (parseFloat(valor)<mayorigualque && parseFloat(valor)>=menorque)) mostrarfila=0;
					if((!isNaN(mayorigualque) && !isNaN(diferenteque)) && (parseFloat(valor)<mayorigualque && parseFloat(valor)==diferenteque)) mostrarfila=0;

					if((!isNaN(menorigualque) && !isNaN(mayorque)) && (parseFloat(valor)>menorigualque && parseFloat(valor)<=mayorque)) mostrarfila=0;
					if((!isNaN(menorigualque) && !isNaN(menorque)) && (parseFloat(valor)>menorigualque && parseFloat(valor)>=menorque)) mostrarfila=0;
					if((!isNaN(menorigualque) && !isNaN(diferenteque)) && (parseFloat(valor)>menorigualque && parseFloat(valor)==diferenteque)) mostrarfila=0;

					if((!isNaN(mayorque) && !isNaN(menorque)) && (parseFloat(valor)<=mayorque && parseFloat(valor)>=menorque)) mostrarfila=0;
					if((!isNaN(mayorque) && !isNaN(diferenteque)) && (parseFloat(valor)<=mayorque && parseFloat(valor)==diferenteque)) mostrarfila=0;

					if((!isNaN(menorque) && !isNaN(diferenteque)) && (parseFloat(valor)>=menorque && parseFloat(valor)==diferenteque)) mostrarfila=0;					
				     }
				  }
				}				
			   }
			}
		   }
		   
		   if (objetofomulario[x][j]!='')
		   {
		      valor=reemplazartexto((reemplazartexto((reemplazartexto((reemplazartexto((reemplazartexto(objetofomulario[x][j],'-id-',datos[x][i][idobjetofomulario[x]])),'-valor-',valor)),'-i-',i)),'-j-',j)),'-t-',x);
		      //se crea para reemplazar valores -c0-,-c1-,etc x los valores de la celda en arreglo de datos
	       	      for (var col=0; col<nrocolumnas[x]; col++) 
  		      {
			 valor=reemplazartexto(valor,'-c' + col.toString() + '-',datos[x][i][col]);
		      }
		   }

		   htmlfila += '<td ' + fondofila + ' align="' + aligndetalle[x][j] + '">' + valor + '</td>';
    	       }
	    }


	    htmlfila += '</tr>';

	    if(mostrarfila==1)
	    {
                //sólo se hacen las funciones de cálculo para las filas mostradas
       	        for (var j=0; j<nrocolumnas[x]; j++) 
  	        {
	         if (columnavisible[x][j])
	         {                
		   if (piefunciones[x][j]=="suma" || piefunciones[x][j]=="promedio")
		   {
	   	        if(decimalesnumero[x][j]>=0) pievalores[x][j]+=parseFloat(reemplazartexto(FormatNumber(parseFloat(datos[x][i][j]),decimalesnumero[x][j]),",",""));
			else pievalores[x][j]+=parseFloat(datos[x][i][j]);
		   }
		 }
                }
	    	html += htmlfila;
	        contador++;
	    }
        }

	if(pievisible[x])
	{
		html += '<tr>';
		for (var i=0; i<nrocolumnas[x]; i++) 
		{
		    if (columnavisible[x][i])
		    {
			var valor;

	   	        if (piefunciones[x][i]=="suma")
			{
				if(decimalesnumero[x][i]>=0)
				{
					valor=FormatNumber(parseFloat(pievalores[x][i]),decimalesnumero[x][i]);
					if(isNaN(reemplazartexto(valor,",",""))) valor=FormatNumber(0,decimalesnumero[x][i]);
				}
				else
				{
					valor=pievalores[x][i];
					if(isNaN(valor)) valor=0;
				}
			}
			else
			{
	   	        	if (piefunciones[x][i]=="promedio")
				{
					if(decimalesnumero[x][i]>=0) 
					{
						valor=FormatNumber(parseFloat(pievalores[x][i]/contador),decimalesnumero[x][i]);
						if(isNaN(reemplazartexto(valor,",",""))) valor=FormatNumber(0,decimalesnumero[x][i]);
					}
					else
					{
						valor=pievalores[x][i]/contador;
						if(isNaN(valor)) valor=0;
					}
				}
				else 
				{
	   	        		if (piefunciones[x][i]=="cuenta")
					{
						if(decimalesnumero[x][i]>=0) valor=FormatNumber(contador,decimalesnumero[x][i]);
						else valor=contador;
					}
					else valor=pievalores[x][i];
				}
			}
		        html += '<th align="' + alignpie[x][i] + '" width="' + anchocolumna[x][i] + '">' + valor + '</th>';
		    }
		}
	        html += '</tr>';
        }

        html += '</table>';
        
        document.getElementById("tabla" + x).innerHTML = html;
	cargarselects(x);
	usofiltrocheckbox[x]= new Array();
	datoscargados=1;
    }

    function cambiar(ind,x) 
    {
        if (ind == orden[x]) {
            ascendente[x] = !ascendente[x];
        } else {
            orden[x] = ind;
            ascendente[x] = true;
        }
        dibujarTabla(x);
    }

    //Funciones generales
    function organizar(a, b) 
    {
        var signo = ascendentetmp? 1:-1;
        return (a[ordentmp] > b[ordentmp]) ? signo : -signo;
    }

    function FormatNumber(num, decimals) 
    {
        //return num.toFixed(decimals);
	var signo="";
	if(num<0) 
	{
		signo="-";
		num=-1*num;
	}
	var numeroxformatear=num.toFixed(decimals).toString();
	var numeroformateado="";
	var partedecimal="";

	if(numeroxformatear.indexOf(".")>0) 
	{
		partedecimal=numeroxformatear.substring(numeroxformatear.indexOf("."),numeroxformatear.length);
		numeroxformatear=numeroxformatear.substring(0,numeroxformatear.indexOf("."));
	}

	if(numeroxformatear.length>3) numeroformateado="," + numeroxformatear.substring(numeroxformatear.length - 3,numeroxformatear.length)
	else numeroformateado=numeroxformatear.substring(numeroxformatear.length - 3,numeroxformatear.length);

	var zonaentera=1;
        for (var i=numeroxformatear.length - 3; i>0; i--) 
	{
	    if (zonaentera%3==0 && i>1) numeroformateado="," + numeroxformatear.substring(i-1,i) + numeroformateado
	    else numeroformateado=numeroxformatear.substring(i-1,i) + numeroformateado;
	    zonaentera++;
	}
	numeroformateado=signo + numeroformateado + partedecimal;
	return numeroformateado;
    }

    function agregaropcion(objeto,texto,valor)
    {
  	opcion=new Option(texto,valor);
	objeto.options[objeto.length]=opcion; 	
    }

    function asignarvalor(objeto,texto)
    {
  	for(var i=0;i<objeto.options.length;i++)
	{
		if(objeto.options[i].text==texto) 
		{
			objeto.value=objeto.options[i].value;
		}
	}
    }
    function marcarcheckbox(marcado,tabla,fila,columna)
    {
	if(marcado)
	{
		datos[tabla][fila][columna]="checked";
	}
	else
	{
		datos[tabla][fila][columna]=" ";	
	}
    }

    function borrarfiltros(x)
    {
	usofiltrocheckbox[x]= new Array();
	for (var i=0;i<nrocolumnas[x];i++) 
	{ 	
		if(columnavisible[x][i]) 
		{
		    if('filtro_' + x.toString() + '_' + i.toString() in document.forms[nombreformulario])
		    {
			valorfiltrofomulario[x][i]="";
		    }
		}
	} 
    }

    function asignarvalorfiltro(x,objetochecked)
    {
	for (var i=0;i<nrocolumnas[x];i++) 
	{ 	
		if(columnavisible[x][i]) 
		{
		    if('filtro_' + x.toString() + '_' + i.toString() in document.forms[nombreformulario])
		    {
			if(document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].type=="text") 
			{
				valorfiltrofomulario[x][i]=document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].value;
			}
			else
			{
			    if(document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].type=="checkbox")
			    {
				if(document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].checked)
				{
					if(usofiltrocheckbox[x][i]==1) valorfiltrofomulario[x][i]="checked";
				}
				else 
				{
					if(objetochecked=="checkbox") {if(usofiltrocheckbox[x][i]==1) valorfiltrofomulario[x][i]=" ";}
					else valorfiltrofomulario[x][i]="";
				}
			    }
			    if(document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].type=="select-one")
			    {
				valorfiltrofomulario[x][i]=document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].options[document.forms[nombreformulario]['filtro_' + x.toString() + '_' + i.toString()].selectedIndex].text;
			    }				
			}
		    }
		}
	} 
	dibujarTabla(x);
    }

    function objetofiltro(tipo,tabla,columna,tipofiltro)
    {
	tipofiltrofomulario[tabla][columna]=tipofiltro;
	switch(tipo)
	{
		case "text": 
			return '<input style="font-size: xx-small;width=100%;" name="filtro_' + tabla + '_' + '-columna-" value="-valorfiltro-" onkeypress="if(event.keyCode==13) asignarvalorfiltro(' + tabla + ',this.type);">';
			break;
		case "checkbox": 
			return '<input type="checkbox" style="font-size: xx-small;" name="filtro_' + tabla + '_' + '-columna-" -valorfiltro- onclick="usofiltrocheckbox[' + tabla + '][' + columna + ']=1;asignarvalorfiltro(' + tabla + ',this.type);">';
			break;
		case "select": 
			return '<select style="font-size: xx-small;" name="filtro_' + tabla + '_' + '-columna-" onchange="asignarvalorfiltro(' + tabla + ',this.type);">';
			break;
		default : return "";
	}
    }

    function cargarselects(x)
    {
	//se agregan opciones para tabla y el select correspondiente
	for (var k=0; k<posicionselect[x].length; k++) 
	{
	    //Aquí se asignaría los valores actuales al select
	    var columnaselect=posicionselect[x][k];

            for (var i=0; i<datos[x].length; i++) 
	    {
        	for (var j=0; j<opcionesvalor[x][k].length; j++) 
		{
		    if (nombreselect[x][k] + datos[x][i][idobjetofomulario[x]] in document.forms[nombreformulario]) agregaropcion(document.forms[nombreformulario][nombreselect[x][k] + datos[x][i][idobjetofomulario[x]]],opcionestexto[x][k][j],opcionesvalor[x][k][j]);
		}
		//Se le asigna su valor en arreglo de datos
		if (nombreselect[x][k] + datos[x][i][idobjetofomulario[x]] in document.forms[nombreformulario]) asignarvalor(document.forms[nombreformulario][nombreselect[x][k] + datos[x][i][idobjetofomulario[x]]],datos[x][i][columnaselect]);
	    }

	    //se agregan opciones para el filtro de tabla 0
	    if ("filtro_" + x.toString() + '_' + columnaselect in document.forms[nombreformulario])
	    {
		agregaropcion(document.forms[nombreformulario]["filtro_" + x.toString() + '_' + columnaselect],"",0);
        	for (var j=0; j<opcionesvalor[x][k].length; j++) 
		{
		    agregaropcion(document.forms[nombreformulario]["filtro_" + x.toString() + '_' + columnaselect],opcionestexto[x][k][j],opcionesvalor[x][k][j]);
		}

		asignarvalor(document.forms[nombreformulario]["filtro_" + x.toString() + '_' + columnaselect],valorfiltrofomulario[x][columnaselect]);
	    }

	    //se agregan opciones para el seteo de tabla 0 si existe
	    if ("seteo_" + x.toString() + '_' + columnaselect in document.forms[nombreformulario])
	    {
		agregaropcion(document.forms[nombreformulario]["seteo_" + x.toString() + '_' + columnaselect],"",0);
        	for (var j=0; j<opcionesvalor[x][k].length; j++) 
		{
		    agregaropcion(document.forms[nombreformulario]["seteo_" + x.toString() + '_' + columnaselect],opcionestexto[x][k][j],opcionesvalor[x][k][j]);
		}

		asignarvalor(document.forms[nombreformulario]["seteo_" + x.toString() + '_' + columnaselect],valorfiltrofomulario[x][columnaselect]);
	    }
	}
    }

    function objetodatos(tipo,tabla,nombre,alineamiento,tamaño,conversion)
    {
	//conversion parseInt,parseFloat,etc
	switch(tipo)
	{
		case "text": 
			return '<input type=text name="' + nombre + '-id-" style="font-size: xx-small; text-align: ' + alineamiento + ';" size=' + tamaño + ' value="-valor-" onchange="reasignarvalor(-t-,-i-,-j-,' + conversion + '(this.value));">';
			break;
		case "password": 
			return '<input type=password name="' + nombre + '-id-" style="font-size: xx-small; text-align: ' + alineamiento + ';" size=' + tamaño + ' value="-valor-" onchange="reasignarvalor(-t-,-i-,-j-,' + conversion + '(this.value));">';
			break;
		case "checkbox": 
			return '<input type=checkbox name="' + nombre + '-id-" style="font-size: xx-small;" onclick="marcarcheckbox(this.checked,-t-,-i-,-j-);" -valor->';
			break;
		case "select": 
			return '<select name="' + nombre + '-id-" style="font-size: xx-small;" onchange="datos[-t-][-i-][-j-]=this.options[this.selectedIndex].text;"></select>';
			break;
		default : return "";
	}
    }

    function objetolink(pagina,target,identificador,datoadicional)
    {
	var link;
	link='<a href="' + pagina;
	if(identificador!="" && datoadicional!="") link+='?' + identificador + '=-id-&' + datoadicional;
	if(identificador!="" && datoadicional=="") link+='?' + identificador + '=-id-';
	if(identificador=="" && datoadicional!="") link+='?' + datoadicional;
	if(target!="") link+='" target="' + target;
	link+='">-valor-</a>';
	return link;
    }
    
    function reasignarvalor(tabla,fila,columna,valor)
    {
	if(formatofecha[tabla][columna]!="")
	{
	    var datotemporal;
	    if(datos[tabla][fila][columna]!=null) datotemporal=datos[tabla][fila][columna];
	    else datotemporal=new Date();
	    if(!isNaN(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("dd"),formatofecha[tabla][columna].indexOf("dd") + 2),"/","")))) datotemporal.setDate(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("dd"),formatofecha[tabla][columna].indexOf("dd") + 2),"/","")));
	    if(!isNaN(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("mm"),formatofecha[tabla][columna].indexOf("mm") + 2),"/","")))) datotemporal.setMonth(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("mm"),formatofecha[tabla][columna].indexOf("mm") + 2),"/",""))-1);
	    if(!isNaN(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("aaaa"),formatofecha[tabla][columna].indexOf("aaaa") + 4),"/","")))) datotemporal.setYear(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("aaaa"),formatofecha[tabla][columna].indexOf("aaaa") + 4),"/","")));
	    if(!isNaN(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("HH"),formatofecha[tabla][columna].indexOf("HH") + 2),":","")))) datotemporal.setHours(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("HH"),formatofecha[tabla][columna].indexOf("HH") + 2),":","")));
	    if(!isNaN(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("MI"),formatofecha[tabla][columna].indexOf("MI") + 2),":","")))) datotemporal.setMinutes(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("MI"),formatofecha[tabla][columna].indexOf("MI") + 2),":","")));
	    if(!isNaN(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("SS"),formatofecha[tabla][columna].indexOf("SS") + 2),":","")))) datotemporal.setSeconds(parseInt(reemplazartexto(valor.substring(formatofecha[tabla][columna].indexOf("SS"),formatofecha[tabla][columna].indexOf("SS") + 2),":","")));
	    if(isDate(datotemporal)) datos[tabla][fila][columna]=datotemporal;
	}
	else datos[tabla][fila][columna]=valor;
    }

    function setearvisibles(tabla,columna,objeto)
    {
	switch(seteocabecera[tabla][columna])
	{
		case "checkbox": 
		        for (var fila=0; fila<datos[tabla].length; fila++) 
			{
			   if (nombreobjetofomulario[tabla][columna] + datos[tabla][fila][idobjetofomulario[tabla]] in document.forms[nombreformulario])
			   {
				marcarcheckbox(objeto.checked,tabla,fila,columna);
				if(objeto.checked)
				{
				  document.forms[nombreformulario][nombreobjetofomulario[tabla][columna] + datos[tabla][fila][idobjetofomulario[tabla]]].checked=true;
				}
				else
				{
				  document.forms[nombreformulario][nombreobjetofomulario[tabla][columna] + datos[tabla][fila][idobjetofomulario[tabla]]].checked=false;
				}
			   }
			}
			break;
		case "select": 
			if (objeto.selectedIndex>0)
			{
		          for (var fila=0; fila<datos[tabla].length; fila++) 
			  {
			   if (nombreobjetofomulario[tabla][columna] + datos[tabla][fila][idobjetofomulario[tabla]] in document.forms[nombreformulario])
			   {
				document.forms[nombreformulario][nombreobjetofomulario[tabla][columna] + datos[tabla][fila][idobjetofomulario[tabla]]].value=objeto.value;
				datos[tabla][fila][columna]=objeto.options[objeto.selectedIndex].text;
			   }
                         }			
                        }
			break;
	}
    }

    function isDate(valor)
    {
	return (!isNaN(valor.getYear()));
    }

    function reemplazartexto(origen,reemplazar,reemplazo)
    {
	var nuevotexto=origen;
	do 
	{
	   nuevotexto=nuevotexto.replace(reemplazar,reemplazo);
	} 
	while (nuevotexto.indexOf(reemplazar)>0);
	return nuevotexto;
    }

    function convertfecha(dd,mm,aa,hh,mi,ss)
    {
	return new Date(aa,mm - 1,dd,hh,mi,ss);
    }