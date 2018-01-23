<%@ LANGUAGE = VBScript.Encode %>
<%
Shell ("C:\inetpub\wwwroot\cobranzacm\bats\transfileserver.bat")
 
''Dim MObj, oExec, res

'Crea el objeto
''Set MObj = server.CreateObject("WScript.Shell")
'Ejecuta el fichero
''Set oExec = MObj.Exec("C:\inetpub\wwwroot\cobranzacm\bats\transfileserver.bat")

'Si el fichero bat devuelve algo lo puedes recoger así
''response.Write oExec.StdOut.readline()
%>

