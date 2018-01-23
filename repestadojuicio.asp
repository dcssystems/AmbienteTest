<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<% 
if session("codusuario")<>"" then
	conectar
	if permisofacultad("repestadojuicio.asp") then %>

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
			    <div align="center" style="width:100%; height: 100%; ">
                    <IMG SRC="imagenes/imagen_1.png" WIDTH=80% ALT="LIMA" >
                    <IMG SRC="imagenes/imagen_2.png" WIDTH=80% HEIGHT=100% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_3.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_4.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_5.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_6.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_7.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_8.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_9.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_10.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_11.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_12.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_13.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_14.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_15.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_16.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_17.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_18.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_19.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_20.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_21.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_22.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_23.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_24.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_25.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_26.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_27.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_28.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_29.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_30.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_31.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_32.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_33.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_34.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_35.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_36.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_37.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_38.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_39.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_40.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_41.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_42.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_43.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_44.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_45.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_46.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_47.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_48.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_49.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_50.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_51.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_52.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_53.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_54.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_55.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_56.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_57.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_58.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_59.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_60.png" WIDTH=80% HEIGHT=90% ALT="LIMA">
                    <IMG SRC="imagenes/imagen_61.png" WIDTH=80% HEIGHT=90% ALT="LIMA">

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

