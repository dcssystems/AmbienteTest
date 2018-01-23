<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
if session("codusuario")<>"" then
		Dim objConn, strFile
		Dim intCampaignRecipientID

		''Response.Buffer = false

		Response.CacheControl = "Private"
		' Selecciona una de las tres opciones siguientes
		Response.Expires = -1441
		Response.Expires = 0
		Response.ExpiresAbsolute = #1/5/2000 12:12:12#
		session.Timeout=600
		Server.ScriptTimeout=999999999

		conectar
		''Capturo tambien ruta de upload
		if obtener("exportados")="" then
			sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload'"
			consultar sql,RS
			RutaFisicaUpload=RS.Fields(1)
			RS.Close
		else
			sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaExportar'"
			consultar sql,RS
			RutaFisicaUpload=RS.Fields(1)
			RS.Close	
		end if
		desconectar	

		RutaFisica=RutaFisicaUpload				
		NombreArchivo=obtener("nomarch")
		''Response.Write RutaFisica
		''Response.Write NombreArchivo
		dim fs
		set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(RutaFisica & "\" & NombreArchivo) then
		    Dim objStream
		    Set objStream = Server.CreateObject("ADODB.Stream")
		    objStream.Type = 1 'adTypeBinary
		    ''objStream.Type = 2 'adTypeText
		    objStream.Open
		    objStream.LoadFromFile(RutaFisica & "\" & NombreArchivo)
		    Response.ContentType = "application/x-unknown"
		    Response.Addheader "Content-Disposition", "attachment; filename=" & NombreArchivo
		    Response.Clear()
		    Response.BinaryWrite objStream.Read
		    objStream.Close
		    Set objStream = Nothing
		else
		  response.write("Archivo no existe.")
		end if
		set fs=nothing   
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	window.open("index.html","_top");
</script>
<%
end if
%>
