<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
conectar
''traemos maxima fecha de gestion en distribución diaria
''si se seleccions fecha de asignacion inicio y fin igual a esta maxima fecha de gestion
''se cruza con la tabla UltClienteDiario y UltContratoDiario para rapidez de reporte de gestion
sql="select max(fechadatos) from Cliente_FileCarta"
consultar sql,RS    
maxfechadatos=rs.fields(0)
RS.Close
sql="select A.codagencia,B.razonsocial as Agencia,A.codoficina,C.descripcion as Oficina,D.codterritorio,D.descripcion as Territorio from Usuario A left outer join Agencia B on A.codagencia=B.codagencia left outer join Oficina C on A.codoficina=C.codoficina left outer join Territorio D on C.codterritorio=D.Codterritorio where A.codusuario = " & session("codusuario")
consultar sql,RS    
codagenciausuario=rs.fields("codagencia")
agenciausuario=rs.fields("agencia")
codoficinausuario=rs.fields("codoficina")
oficinausuario=rs.fields("oficina")
codterritoriousuario=rs.fields("codterritorio")
territoriousuario=rs.fields("territorio")    
rs.close    
actualizarlista=obtener("actualizarlista")
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
if permisofacultad("admcartas.asp") then
contrato=obtener("contrato1")
codigocentral=obtener("codigocentral1")
codgestor=obtener("codgestor1")    
codtipodocumento=obtener("codtipodocumento1")
numdocumento=obtener("numdocumento1")
segmento=obtener("segmento1")
codproducto=obtener("codproducto1")
codmarca=obtener("codmarca1")
clasificacion=obtener("clasificacion1")    
if obtener("actualizarlista")<>"" then
            codterritorio=obtener("codterritorio1")
            if codterritorio="" then
           codoficina=""
            else
           codoficina=obtener("codoficina1")
            end if    
        else
            codterritorio=obtener("codterritorio")
            if codterritorio="" then
           codoficina=""
            else
           codoficina=obtener("codoficina")
            end if           
        end if    
estado=obtener("estado1")
fechadatosini=obtener("fechadatosini1")
fechadatosfin=obtener("fechadatosfin1")
diaatrasoini=obtener("diaatrasoini1")
diaatrasofin=obtener("diaatrasofin1")
if not IsDate(fechadatosini) then
   fechadatosini=CStr(maxfechadatos)
end if
if not IsDate(fechadatosfin) then
   fechadatosfin=CStr(maxfechadatos)
end if    
       
%>
<html>
<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
<head>
<title>Cartas a Clientes</title>
<script language=javascript src="scripts/TablaDinamica.js"></script>
<script language=javascript>
    var ventanaverficha;
    function inicio() {
        dibujarTabla(0);
    }
    function modificar(fdcodcen) {
        ventanaverficha = window.open("verficha.asp?vistapadre=" + window.name + "&paginapadre=admcartas.asp&fdcodcen=" + fdcodcen, "VerFicha" + fdcodcen, "scrollbars=yes,scrolling=yes,top=" + ((screen.height) / 2 - 300) + ",height=600,width=" + (screen.width / 2 + 300) + ",left=" + (screen.width / 2 - 475) + ",resizable=yes");
        ventanaverficha.focus();
    }
    function actualizar() {
        document.formula.actualizarlista.value = 1;
        document.formula.submit();
    }
    function exportar() {
        if (document.formula.buscando.value == "") {
            document.formula.expimp.value = 1;
            document.formula.submit();
        }
    }
    function imprimir() {
        window.open("impusuarios.asp", "ImpUsuarios", "scrollbars=yes,scrolling=yes,top=0,height=200,width=200,left=0,resizable=yes");
    }
    function buscar() {
        if (document.formula.buscando.value == "") {
            document.formula.buscando.value = "OK";
            document.formula.nuevabusqueda.value = 1;
            document.formula.actualizarlista.value = 1;
            document.formula.pag.value = 1;
            document.formula.contrato1.value = document.formula.contrato.value;
            document.formula.codigocentral1.value = document.formula.codigocentral.value;
            document.formula.codgestor1.value = document.formula.codgestor.value;
            document.formula.codtipodocumento1.value = document.formula.codtipodocumento.value;
            document.formula.numdocumento1.value = document.formula.numdocumento.value;
            document.formula.segmento1.value = document.formula.segmento.value;
            document.formula.codproducto1.value = document.formula.codproducto.value;
            document.formula.codmarca1.value = document.formula.codmarca.value;
            document.formula.clasificacion1.value = document.formula.clasificacion.value;
            document.formula.codterritorio1.value = document.formula.codterritorio.value;
            document.formula.codoficina1.value = document.formula.codoficina.value;
            document.formula.estado1.value = document.formula.estado.value;
            document.formula.fechadatosini1.value = document.formula.fechadatosini.value;
            document.formula.fechadatosfin1.value = document.formula.fechadatosfin.value;
            document.formula.diaatrasoini1.value = document.formula.diaatrasoini.value;
            document.formula.diaatrasofin1.value = document.formula.diaatrasofin.value;
            document.formula.submit();
        }
    }
    function filtrar() {
        if (filtrardatos[0] == 1) {
            filtrardatos[0] = 0;
            dibujarTabla(0);
        }
        else {
            filtrardatos[0] = 1;
            dibujarTabla(0);
        }
    }
    function mostrarpag(pagina) {
        if (document.formula.buscando.value == "") {
            document.formula.buscando.value = "OK";
            document.formula.pag.value = pagina;
            document.formula.submit();
        }
    }
</script>
<style>
A {
FONT-SIZE: 10px; COLOR: #00529B; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
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
font-size:10px;
font-family:Arial;
cursor:hand;
}
</style>

 
<script language=javascript>
    function actualizaterritorio() {
        if (document.formula.buscando.value == "") {
            document.formula.actualizarlista.value = "";
            document.formula.pag.value = 1;
            document.formula.submit();
        }
    }
</script>
<link rel="stylesheet" type="text/css" media="all" href="scripts/calendar-blue2.css" title="blue" />
<script type="text/javascript" src="scripts/calendar.js"></script>
<script type="text/javascript" src="scripts/calendar-es.js"></script>
<script type="text/javascript">
    function selected(cal, date) {
        cal.sel.value = date;
        if (cal.dateClicked && (cal.sel.id == "sel1" || cal.sel.id == "sel3"))
            cal.callCloseHandler();
    }
    function closeHandler(cal) {
        cal.hide();
        _dynarch_popupCalendar = null;
    }
    function showCalendar(id, format, showsTime, showsOtherMonths) {
        var el = document.getElementById(id);
        if (_dynarch_popupCalendar != null) {
            _dynarch_popupCalendar.hide();
        } else {
            var cal = new Calendar(1, null, selected, closeHandler);
            if (typeof showsTime == "string") {
                cal.showsTime = true;
                cal.time24 = (showsTime == "24");
            }
            if (showsOtherMonths) {
                cal.showsOtherMonths = true;
            }
            _dynarch_popupCalendar = cal;
            cal.setRange(1900, 2070);
            cal.create();
        }
        _dynarch_popupCalendar.setDateFormat(format);
        _dynarch_popupCalendar.parseDate(el.value);
        _dynarch_popupCalendar.sel = el;

        _dynarch_popupCalendar.showAtElement(el.nextSibling, "Br");
        return false;
    }
</script>

</head>
<%Response.Flush()%>
<%if actualizarlista<>"" then%>
<script language=javascript>
rutaimgcab="imagenes/"; 
 //Configuración general de datos de tabla 0
   tabla=0;
   orden[tabla]=0;
   ascendente[tabla]=true;
   nrocolumnas[tabla]=13;
   fondovariable[tabla]='bgcolor=#f5f5f5';
   anchotabla[tabla]='100%';
   botonfiltro[tabla] = false;
   botonactualizar[tabla] = false;
   botonagregar[tabla] = false;
paddingtabla[tabla] = '2';
spacingtabla[tabla] = '1';       
   cabecera[tabla] = new Array('FDCli','Fecha','Código','Nombre','Documento','Marca','Segmento','Banca','Territorio','Oficina','D','Clasificación','Estado');
   identificadorfilas[tabla]="fila";
   pievisible[tabla]=true;
   columnavisible[tabla] = new Array( false, true, true,true,true,true, true, true,true,true,true, true, true);
   anchocolumna[tabla] =  new Array( '' , '1%' ,  '1%',  '4%',  '2%','1%' , '1%' ,  '2%',  '3%',  '3%','1%' , '2%' ,  '1%');
   aligncabecera[tabla] = new Array('left','left','left','left','left','left','left','left','left','left','left','left','left');
   aligndetalle[tabla] = new Array('left','left','left','left','left','left','left','left','left','left','left','left','left');
   alignpie[tabla] =     new Array('left','left','left','left','left','left','left','left','left','left','left','left','left');
   decimalesnumero[tabla] = new Array(-1 ,-1   ,-1 ,-1 ,-1,-1 ,-1   ,-1 ,-1 ,-1 ,0  ,-1 ,-1);
   formatofecha[tabla] =   new Array(''  ,'dd/mm/aaaa'  ,'' ,'','',''  ,''  ,'' ,'','',''  ,''  ,'' );


   //Se escriben condiciones de datos administrados "objetos formulario"
   idobjetofomulario[tabla]=0; //columna 1 trae el id de objetos x administrar ejm. zona1543 = 'zona' + idpedido (datos[0][fila][idobjetofomulario[0]])
   objetofomulario[tabla] = new Array();
objetofomulario[tabla][0]='<input type=hidden name=fdcli-id- value=-c0->';
objetofomulario[tabla][1]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][2]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][3]='<a href=javascript:modificar("-c0-");>-valor-</a>';    
objetofomulario[tabla][4]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][5]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][6]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][7]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][8]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][9]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][10]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][11]='<a href=javascript:modificar("-c0-");>-valor-</a>';
objetofomulario[tabla][12]='<a href=javascript:modificar("-c0-");>-valor-</a>';
   filtrardatos[tabla]=0; //define si carga auto el filtro
   filtrofomulario[tabla] = new Array();
   tipofiltrofomulario[tabla] = new Array();
       filtrofomulario[tabla][0]='';
filtrofomulario[tabla][1]='';
filtrofomulario[tabla][2]='';
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
   //Se escribe el conjunto de datos de tabla 0
   datos[tabla]=new Array();
<%
filtrobuscador = " where A.fechadatos between convert(datetime,'" & mid(fechadatosini,7,4) & mid(fechadatosini,4,2) & mid(fechadatosini,1,2) & "') and convert(datetime,'" & mid(fechadatosfin ,7,4) & mid(fechadatosfin,4,2) & mid(fechadatosfin,1,2) & "')"
if contrato<>"" then
filtrobuscador = filtrobuscador & " and A.codigocentral in (select codigocentral from Contrato_FileCarta where fechadatos=A.fechadatos and contrato='" & contrato & "')"
end if
if codigocentral<>"" then
filtrobuscador = filtrobuscador & " and A.codigocentral='" & codigocentral & "'"
end if    
if codgestor<>"" then
filtrobuscador = filtrobuscador & " and B.codgestor=" & codgestor
end if    
if codtipodocumento<>"" then
filtrobuscador = filtrobuscador & " and A.tipodocumento='" & codtipodocumento & "'"
end if    
if numdocumento<>"" then
filtrobuscador = filtrobuscador & " and A.numdocumento='" & numdocumento & "'"
end if    
if segmento<>"" then
filtrobuscador = filtrobuscador & " and A.segmento_riesgo='" & segmento & "'"
end if    
if codproducto<>"" then
filtrobuscador = filtrobuscador & " and A.codigocentral in (select codigocentral from Contrato_FileCarta where codproducto='" & codproducto & "' and fechadatos=A.fechadatos)"
end if    
if codmarca<>"" then
filtrobuscador = filtrobuscador & " and A.marca=(select descripcion from Marca where codmarca=" & codmarca & ")"
end if    
if clasificacion<>"" then
filtrobuscador = filtrobuscador & " and A.clasificacion=" & clasificacion
end if    
if codterritorio<>"" then
filtrobuscador = filtrobuscador & " and A.codterritorio='" & codterritorio & "'"
end if    
if codoficina<>"" then
filtrobuscador = filtrobuscador & " and A.codoficina='" & codoficina & "'"
end if    
if estado<>"" then
filtrobuscador = filtrobuscador & " and B.estado=" & estado
end if    
if diaatrasoini<>"" then
filtrobuscador = filtrobuscador & " and A.maxdias>=" & diaatrasoini
end if    
if diaatrasofin<>"" then
filtrobuscador = filtrobuscador & " and A.maxdias<=" & diaatrasofin
end if    
if filtrobuscador<>"" then
filtrobuscador1=mid(filtrobuscador,7,len(filtrobuscador)) & " and "
end if    
''response.write filtrobuscador
contadortotal=0
sql="select count(*) from Cliente_FileCarta A " & filtrobuscador
''response.write sql
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
   sql="select top " & cantidadxpagina & " convert(varchar,A.fechadatos,112) + A.codigocentral as FDCli,(select descripcion from EstadoCarta where codestado=A.estado) as estado,(select descripcion from Clasificacion where codclasificacion=A.clasificacion) as clasifica) " & filtrobuscador1 & " convert(varchar,A.fechadatos,112) + A.codigocentral not in (select top " & topnovisible & " convert(varchar,A.fechadatos,112) + A.codigocentral  from Cliente_FichaVisita A inner join Detalle_FichaVisita B on A.codigocentral=B.codigocentral and A.fechadatos=B.fechadatos " & filtrobuscador & " order by convert(varchar,A.fechadatos,112) + A.codigocentral) order by convert(varchar,A.fechadatos,112) + A.codigocentral"
else
   sql="select top " & cantidadxpagina & " convert(varchar,A.fechadatos,112) + A.codigocentral as FDCli,(select descripcion from EstadoCarta where codestado=A.estado) as estado,(select descripcion from Clasificacion where codclasificacion=A.clasificacion) as clasifica) " & filtrobuscador & " order by convert(varchar,A.fechadatos,112) + A.codigocentral"
end if
''response.write sql
consultar sql,RS
contador=0
''FDCli','Fecha','Código','Nombre','Tipo','Documento','Marca','Segmento','Banca','Territorio','Oficina','D','Clasificación','Estado
Do while not RS.EOF
%>
datos[tabla][<%=contador%>] = new Array();
datos[tabla][<%=contador%>][0]='<%=RS.Fields("FDCli")%>';
   datos[tabla][<%=contador%>][1]=<%if not IsNull(RS.Fields("fechadatos")) then%>new Date(<%=Year(RS.Fields("fechadatos"))%>,<%=Month(RS.Fields("fechadatos"))-1%>,<%=Day(RS.Fields("fechadatos"))%>,0,0,0)<%else%>null<%end if%>;
       datos[tabla][<%=contador%>][2]='<%=RS.Fields("codigocentral")%>';
datos[tabla][<%=contador%>][3]='<%if len(trim(replace(RS.Fields("nombres"),"'","´")))<=27 then%><%=trim(replace(RS.Fields("nombres"),"'","´"))%><%else%><%=mid(trim(replace(RS.Fields("nombres"),"'","´")),1,27) & "..."%><%end if%>';
datos[tabla][<%=contador%>][4]='<%=RS.Fields("tipodocumento") + " - " + RS.Fields("numdocumento")%>';
datos[tabla][<%=contador%>][5]='<%=RS.Fields("marca")%>';
datos[tabla][<%=contador%>][6]='<%=RS.Fields("segmento_riesgo")%>';
datos[tabla][<%=contador%>][7]='<%=RS.Fields("banca")%>';
datos[tabla][<%=contador%>][8]='<%=RS.Fields("codterritorio") + " - " + RS.Fields("territorio")%>';
datos[tabla][<%=contador%>][9]='<%=RS.Fields("codoficina") + " - " + RS.Fields("oficina")%>';
datos[tabla][<%=contador%>][10]=<%=RS.Fields("maxdias")%>;
datos[tabla][<%=contador%>][11]='<%=RS.Fields("clasifica")%>';
datos[tabla][<%=contador%>][12]='<%=RS.Fields("estado")%>';
<%
contador=contador + 1
RS.MoveNext 
Loop 
RS.Close
%>
   
   //datos del pie si fuera visible
   pievalores[tabla] = new Array('&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;','&nbsp;');
   piefunciones[tabla] = new Array('','','','','','','','','','','','',''); 


   //Se escriben las opciones para los selects que contenga
   posicionselect[tabla]=new Array();
   nombreselect[tabla]=new Array();
   opcionesvalor[tabla]=new Array();
   opcionestexto[tabla]=new Array();
   //Finaliza configuracion de tabla 0
   
funcionactualiza[tabla]='document.formula.actualizarlista.value=1;document.formula.submit();';
   funcionagrega[tabla]='agregar();';

</script> 
<%end if%>
<body topmargin=0 leftmargin=0 bgcolor="#FFFFFF">
<form name=formula method=post>
<table border=0 cellspacing=0 cellpadding=2 width=100%>
  <tr bgcolor="#007DC5">
    <td colspan="17" align="center" height="22">
      <font size=2 face=Arial color="#FFFFFF"><b>Gestión de Cartas </b></font>
    </td>
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>N°&nbsp;Contrato:</b></font></td>
    <td colspan="4"><input name="contrato" type=text maxlength=18 value="<%=contrato%>" style="font-size: x-small; width: 250px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Código&nbsp;Cliente:</b></font></td>
    <td colspan="4"><input name="codigocentral" type=text maxlength=8 value="<%=codigocentral%>" style="font-size: x-small; width: 250px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Gestor:</b></font></td>
    <td colspan="4">
<select name="codgestor" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Gestor</option>
<%
sql = "select A.codusuario,A.nombres,A.apepaterno,A.apematerno from Usuario A inner join TipoUsuario B on A.codtipousuario=B.codtipousuario and B.descripcion='Gestor' and A.activo=1"
consultar sql,RS
Do While Not  RS.EOF
nombregestor=iif(IsNull(RS.Fields("nombres")),"",RS.Fields("nombres")) & ", " & iif(IsNull(RS.Fields("apepaterno")),"",RS.Fields("apepaterno")) & " " & iif(IsNull(RS.Fields("apematerno")),"",RS.Fields("apematerno"))
%>
<option value="<%=RS.Fields("codusuario")%>" <%if codgestor<>"" then%><%if RS.fields("codusuario")=int(codgestor) then%> selected<%end if%><%else%><%if actualizarlista="" and RS.fields("codusuario")=int(session("codusuario")) then%> selected<%end if%><%end if%>><%=nombregestor%></option>
<%
RS.MoveNext
loop
RS.Close
%>    
</select>    
</td>      
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Tipo&nbsp;Documento:</b></font></td>
    <td colspan="4">
<select name="codtipodocumento" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Tipo Documento</option>
<%
sql = "select codtipodocumento,descripcion from TipoDocumento where activo=1 and codtipodocumento in (select distinct tipodocumento from Cliente_FileCarta) order by descripcion"
consultar sql,RS
Do While Not RS.EOF
%>
<option value="<%=RS.Fields("codtipodocumento")%>" <%if codtipodocumento<>"" then%><% if RS.fields("codtipodocumento")=codtipodocumento then%> selected<%end if%><%end if%>><%=RS.Fields("CodTipoDocumento") & " - " & RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
%>    
</select>
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>N°&nbsp;Documento:</b></font></td>
    <td colspan="4"><input name="numdocumento" type=text maxlength=15 value="<%=numdocumento%>" style="font-size: x-small; width: 250px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Segmento:</b></font></td>
    <td colspan="4">
    <select name="segmento" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Segmento</option>
<%
sql = "select distinct segmento_riesgo from Cliente_FileCarta order by segmento_riesgo"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("segmento_riesgo")%>" <%if segmento<>"" then%><%if RS.fields("segmento_riesgo")=segmento then%> selected<%end if%><%end if%>><%=RS.Fields("segmento_riesgo")%></option>
<%
RS.MoveNext
loop
RS.Close
%>
</select>
</td>     
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Producto</b></font></td>
    <td colspan="4">
        <select name="codproducto" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Producto</option>
<%
sql = "select codproducto, descripcion from producto where activo=1 and codproducto in (select distinct codproducto from Contrato_FileCarta) order by codproducto"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("codproducto")%>" <% if codproducto<>"" then%><% if RS.fields("codproducto")=codproducto then%> selected<%end if%><%end if%>><%=RS.Fields("codproducto") & " - " & RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
%>
</select>
</td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Marca:</b></font></td>
    <td colspan="4">
    <select name="codmarca" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Marca</option>
<%
sql = "select codmarca,descripcion from Marca where activo=1 and descripcion in (select distinct marca from Cliente_FileCarta) order by codmarca"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("codmarca")%>" <% if codmarca<>"" then%><% if RS.fields("codmarca")=int(codmarca) then%> selected<%end if%><%end if%>><%=RS.Fields("codmarca") & " - " & RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
%>    
</select>
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>    
    <td align="right"><font size=1 face=Arial color=#00529B><b>Clasificación:</b></font></td>
    <td colspan="4">
        <select name="clasificacion" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Clasificación</option>
<%
sql = "select codclasificacion,descripcion from Clasificacion where activo=1 order by codclasificacion"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("codclasificacion")%>" <%if clasificacion<>"" then%><% if int(clasificacion)=RS.Fields("codclasificacion") then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
%>    
</select>    
</select>
</td>        
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Territorio:</b></font></td>
    <td colspan="4">
<%if codoficinausuario<>"" then%>
<font size=1 face=Arial color=#00529B><%=TerritorioUsuario%></font>
<%else%>    
<select name="codterritorio" onchange="actualizaterritorio()" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Territorio</option>
<%
sql = "select distinct codterritorio, territorio as descripcion from Cliente_FileCarta order by codterritorio"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("codterritorio")%>" <%if codterritorio<>"" then%><% if RS.fields("codterritorio")=codterritorio then%> selected<%end if%><%end if%>><%=RS.Fields("codterritorio") & " - " & RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
%>
</select>
<%end if%>
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Oficina:</b></font></td>
    <td colspan="4">
<%if codoficinausuario<>"" then%>
<font size=1 face=Arial color=#00529B><%=codoficinausuario & " - " & OficinaUsuario%></font><input type=hidden name="codoficina" value="<%=codoficinausuario%>">
<%else%>        
<select name = "codoficina" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Oficina</option>
<%
if codterritorio<>"" then
sql = "select distinct codoficina, oficina as descripcion from Cliente_FileCarta where codterritorio = " & codterritorio & " order by codoficina"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("codoficina")%>" <% if codoficina<>"" then%><% if RS.fields("codoficina")=codoficina then%> selected<%end if%><%end if%>><%=RS.Fields("CodOficina") & " - " & RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
end if
%>
</select>
<%end if%> 
    </td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Estado:</b></font></td>
    <td colspan="4">
        <select name="estado" style="font-size: x-small; width: 250px;">
<option value="">Seleccione Estado</option>
<%
sql = "select codestado,descripcion from EstadoCarta where activo=1 order by codestado"
consultar sql,RS
Do While Not  RS.EOF
%>
<option value="<%=RS.Fields("codestado")%>" <%if estado<>"" then%><% if int(estado)=RS.Fields("codestado") then%> selected<%end if%><%end if%>><%=RS.Fields("Descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close
%>    
</select>    
</td>    
  </tr>
  <tr bgcolor="#E9F8FE">
    <td align="right"><font size=1 face=Arial color=#00529B><b>Fecha&nbsp;de&nbsp;Datos:</b></font></td>
    <td width="40" align="right"><font size=1 face=Arial color=#00529B><b>Del</b></font></td>
    <td width="78"><input name="fechadatosini" id="sel0" readonly  type=text maxlength=10 size=10 value="<%if IsDate(fechadatosini) then%><%=fechadatosini%><%else%><%=maxfechadatos%><%end if%>"  style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel0', '%d/%m/%Y');"></td>
    <td width="40" align="right"><font size=1 face=Arial color=#00529B><b>al</b></font></td>
    <td>
<input name="fechadatosfin" type=text maxlength=10 id="sel3" readonly value="<%if IsDate(fechadatosfin) then%><%=fechadatosfin%><%else%><%=maxfechadatos%><%end if%>" style="font-size: x-small; width: 60px;"><input type="image" style="vertical-align: bottom;" src="imagenes/minicalendar.png" border=0 onclick="return showCalendar('sel3', '%d/%m/%Y');">
<!--<input type="text" name="date3" id="sel3" size="10"><input type="image" value=" ... " onclick="return showCalendar('sel3', '%d/%m/%Y');">-->
</td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td align="right"><font size=1 face=Arial color=#00529B><b>Días&nbsp;de&nbsp;Atraso:</b></font></td>
    <td width="40" align="right"><font size=1 face=Arial color=#00529B><b>Del</b></font></td>
    <td width="78"><input name="diaatrasoini" type=text maxlength=10 value="<%=diaatrasoini%>" style="text-align: right;font-size: x-small; width: 50px;"></td>
    <td width="40" align="right"><font size=1 face=Arial color=#00529B><b>al</b></font></td>
    <td><input name="diaatrasofin" type=text maxlength=10 value="<%=diaatrasofin%>" style="text-align: right;font-size: x-small; width: 50px;"></td>
    <td><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
    <td colspan=5><font size=1 face=Arial color=#00529B>&nbsp;</font></td>
  </tr>
  <tr>
  <td colspan="17">
<%if contador=0 then%>
<table width=100% cellpadding=4 cellspacing=0>    
<tr>
<td bgcolor="#F5F5F5"><font size=1 face=Arial color=#00529B><b>Cartas (0) - No hay registros.</b></font>&nbsp;<a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a></td>
</tr>
</table>
<%else    
%>
<table width=100% cellpadding=4 cellspacing=0 border=0>    
<tr>
<td bgcolor="#F5F5F5" align=left><font size=1 face=Arial color=#00529B><b>Cartas (<%=contadortotal%>)&nbsp;&nbsp;<a href="javascript:buscar();"><img src="imagenes/buscar.gif" border=0 alt="Buscar" title="Buscar" align=middle></a>&nbsp;&nbsp;<a href="javascript:exportar();"><img src="imagenes/excel.gif" border=0 alt="Exportar a Excel" title="Exportar a Excel" align=middle></a><!--&nbsp;&nbsp;<a href="javascript:imprimir();"><img src="imagenes/imprimir.gif" border=0 alt="Imprimir" title="Imprimir" align=middle></a>--><%if expimp="1" then%>&nbsp;&nbsp;<a href='<%=RutaWebExportar%>/UserExport<%=session("codusuario")%>.xls?time=<%=tiempoexport%>','_self'><img src="imagenes/descargar.gif" border=0 alt="Descargar Excel" title="Descargar Excel" align=middle></a><%end if%></b></font></td>
<td bgcolor="#F5F5F5" align=right width=180><font size=1 face=Arial color=#00529B>Pág.&nbsp;<%if bloqueactual>1 then%><a href="javascript:mostrarpag(1);"><<</a>&nbsp;<%end if%><%if bloqueactual>1 then%><a href="javascript:mostrarpag(<%=(bloqueactual-1)*paginasxbloque%>);"><</a>&nbsp;<%end if%><%if pagmax>bloqueactual*paginasxbloque then valorhasta=bloqueactual*paginasxbloque else valorhasta=pagmax end if%><%for i=(bloqueactual - 1)*paginasxbloque + 1 to valorhasta%><%if pag=i then%>[<%else%><a href="javascript:mostrarpag(<%=i%>);"><%end if%><%=i%><%if pag=i then%>]<%else%></a><%end if%>&nbsp;<%next%><%if pagmax>bloqueactual*paginasxbloque then%><a href="javascript:mostrarpag(<%=(bloqueactual)*paginasxbloque + 1%>);">></a>&nbsp;<%end if%><%if bloqueactual<bloquemax then%><a href="javascript:mostrarpag(<%=pagmax%>);">>></a>&nbsp;<%end if%></font></td>
</tr>    
</table>
<div id="tabla0"> 
</div>
<%end if%>
        <input type="hidden" name="buscando" value="">    
<input type="hidden" name="actualizarlista" value="<%=actualizarlista%>">
<input type="hidden" name="nuevabusqueda" value="">
<input type="hidden" name="expimp" value="">    
<input type="hidden" name="pag" value="<%=pag%>">
<input type="hidden" name="contrato1" value="<%=obtener("contrato1")%>">
<input type="hidden" name="codigocentral1" value="<%=obtener("codigocentral1")%>">
<input type="hidden" name="codgestor1" value="<%=obtener("codgestor1")%>">
<input type="hidden" name="codtipodocumento1" value="<%=obtener("codtipodocumento1")%>">
<input type="hidden" name="numdocumento1" value="<%=obtener("numdocumento1")%>">
<input type="hidden" name="segmento1" value="<%=obtener("segmento1")%>">
<input type="hidden" name="codproducto1" value="<%=obtener("codproducto1")%>">    
<input type="hidden" name="codmarca1" value="<%=obtener("codmarca1")%>">    
<input type="hidden" name="clasificacion1" value="<%=obtener("clasificacion1")%>">    
        <input type="hidden" name="codterritorio1" value="<%=obtener("codterritorio1")%>">
<input type="hidden" name="codoficina1" value="<%=obtener("codoficina1")%>">
<input type="hidden" name="estado1" value="<%=obtener("estado1")%>">
<input type="hidden" name="fechadatosini1" value="<%=obtener("fechadatosini1")%>">
<input type="hidden" name="fechadatosfin1" value="<%=obtener("fechadatosfin1")%>">    
<input type="hidden" name="diaatrasoini1" value="<%=obtener("diaatrasoini1")%>">
<input type="hidden" name="diaatrasofin1" value="<%=obtener("diaatrasofin1")%>">
<%if actualizarlista<>"" then%>
<script language="javascript">
    inicio();
</script>    
<%end if%>
<!--cargando--><script language=javascript>                   document.getElementById("imgloading").style.display = "none";</script>    
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
''FDCli','Fecha','Código','Nombre','Tipo','Documento','Marca','Segmento','Banca','Territorio','Oficina','D','Clasificación','Estado
consulta_exp="select 'Fecha','Código','Nombre','Documento','Marca','Segmento','Banca','Territorio','Oficina','D','Clasificación','Estado'"
sql="EXEC SP_EXPEXCEL '" & replace(consulta_exp,"'","''''") & "','" & conn_server & "','" & conn_uid & "','" & conn_pwd & "','" & RutaFisicaExportar & "\temp1_" & session("codusuario") & ".txt'"
conn.execute sql
''Segundo Detalle en temp2_(user).txt
consulta_exp="select convert(varchar,A.fechadatos,103),char(39) + A.codigocentral,A.nombres,A.TipoDocumento + ' - ' + A.NumDocumento,A.Marca,A.Segmento_Riesgo,A.Banca,A.codterritorio + ' - ' + A.territorio,A.codoficina + ' - ' + A.oficina,A.maxdias,(select descripcion from EstadoCarta where codestado=A.estado) as estado,(select descripcion from Clasificacion where codclasificacion=A.clasificacion) as clasifica) " & _
"from CobranzaCM.dbo.Cliente_FileCarta A " & replace(filtrobuscador,"from ","from CobranzaCM.dbo.") & " order by convert(varchar,A.fechadatos,112) + A.codigocentral"
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
  </td>
  </tr>   
</table>
</form>
</body>
<!--cargando--><script language=javascript>                   document.getElementById("imgloading").style.display = "none";</script>
</html>
<%
else
%>
<script language="javascript">
    alert("Ud. No tiene autorización para este proceso.");
    window.open("userexpira.asp", "_top");
</script>
<%    
end if
desconectar
else
%>
<script language="javascript">
    alert("Tiempo Expirado");
    window.open("index.html", "_top");
</script>
<%
end if
%>
