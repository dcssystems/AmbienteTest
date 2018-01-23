<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<%
if session("codusuario")<>"" then
	conectar
		Dim objConn, strFile
		Dim intCampaignRecipientID

		Response.Buffer = false

		Response.CacheControl = "Private"
		' Selecciona una de las tres opciones siguientes
		Response.Expires = -1441
		Response.Expires = 0
		Response.ExpiresAbsolute = #1/5/2000 12:12:12#
		session.Timeout=600
		Server.ScriptTimeout=999999999

		sql="select valortexto1 from parametro where descripcion='RutaFisicaExportar'"
		consultar sql,RS
		RutaFisicaExportar=RS.Fields(0)
		RS.Close

		NombreArchivo=request("archivo")

		dim fs
		set fs=Server.CreateObject("Scripting.FileSystemObject")
		if fs.FileExists(RutaFisicaExportar & "\" & NombreArchivo) then
		    Dim objStream
		    Set objStream = Server.CreateObject("ADODB.Stream")
		    objStream.Type = 1 'adTypeBinary
		    ''objStream.Type = 2 'adTypeText
		    objStream.Open
		    objStream.LoadFromFile(RutaFisicaExportar & "\" & NombreArchivo)
		    Response.ContentType = "application/x-unknown"
		    Response.Addheader "Content-Disposition", "attachment; filename=" & NombreArchivo
		    Response.BinaryWrite objStream.Read
		    objStream.Close
		    Set objStream = Nothing
		else
		  response.write("Archivo no existe.")
		end if
		set fs=nothing
	desconectar
else
%>
<script language="javascript">
	alert("Tiempo Expirado");
	window.open("index.html","sistema");
	window.close();
</script>
<%
end if		    
%>
