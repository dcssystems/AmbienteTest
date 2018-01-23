<%@ LANGUAGE = VBScript.Encode %>
<%
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

RutaFisica="\\s7729w32\xcobranzas\ExportadosWeb\"
NombreArchivo="UserExport.xls"

RutaFisica="D:\cobranzacm\fileserver"
NombreArchivo="AVALES20140703.txt"

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
    Response.BinaryWrite objStream.Read
    objStream.Close
    Set objStream = Nothing
else
  response.write("Archivo no existe.")
end if
set fs=nothing
    
    

%>
