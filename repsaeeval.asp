<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repsaeeval.asp") then %>

	    <!--Ojo esta ventana siempre es flotante-->
		<html>
		<!--cargando--><%Response.Flush()%><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading" style="margin-left: 50px;margin-top:50px;"><%Response.Flush()%>
			<title>Informe de Evaluación Mensual por Especialista</title>
			
			<style>
			A {
				FONT-SIZE: 12px; COLOR: #483d8b; FONT-FAMILY:"Arial"; TEXT-DECORATION: none
			}
			A:visited {
				TEXT-DECORATION: none; COLOR: #483d8b;
			}
			A:hover {
				COLOR: #483d8b; FONT-FACE:"Arial"; TEXT-DECORATION: none
			}			
            #encabezado{border:0}
            #encabezado th{border-width:1px}
            #datos{border:0}
            #datos td{border-width:1px}			
			</style>
			    <div style="width:100%; height: 100%; overflow-x: scroll;">
                    <IMG SRC="imagenes/LIMA.png" WIDTH=180% ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA1.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA2.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA3.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA4.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA5.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA6.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA7.png" WIDTH=180%  ALT="LIMA">
                    <IMG SRC="imagenes/PROVINCIA8.png" WIDTH=180%  ALT="LIMA">
                </div>
		    
		</html>	
		
		<!--cargando--><script language=javascript>document.getElementById("imgloading").style.display = "none";</script>	
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
	//window.open("index.html","_top");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if
%>

