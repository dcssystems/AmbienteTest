<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
			'***************************************
			' File:	  Upload.asp
			' Author: Jacob "Beezle" Gilley
			' Email:  avis7@airmail.net
			' Date:   12/07/2000
			' Comments: The code for the Upload, CByteString, 
			'			CWideString	subroutines was originally 
			'			written by Philippe Collignon...or so 
			'			he claims. Also, I am not responsible
			'			for any ill effects this script may
			'			cause and provide this script "AS IS".
			'			Enjoy!
			'****************************************

			Class FileUploader
				Public  Files
				Private mcolFormElem

				Private Sub Class_Initialize()
					Set Files = Server.CreateObject("Scripting.Dictionary")
					Set mcolFormElem = Server.CreateObject("Scripting.Dictionary")
				End Sub
				
				Private Sub Class_Terminate()
					If IsObject(Files) Then
						Files.RemoveAll()
						Set Files = Nothing
					End If
					If IsObject(mcolFormElem) Then
						mcolFormElem.RemoveAll()
						Set mcolFormElem = Nothing
					End If
				End Sub

				Public Property Get Form(sIndex)
					Form = ""
					If mcolFormElem.Exists(LCase(sIndex)) Then Form = mcolFormElem.Item(LCase(sIndex))
				End Property

				Public Default Sub Upload()
					Dim biData, sInputName
					Dim nPosBegin, nPosEnd, nPos, vDataBounds, nDataBoundPos
					Dim nPosFile, nPosBound

					biData = Request.BinaryRead(Request.TotalBytes)
					nPosBegin = 1
					nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
					
					If (nPosEnd-nPosBegin) <= 0 Then Exit Sub
					 
					vDataBounds = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
					nDataBoundPos = InstrB(1, biData, vDataBounds)
					
					Do Until nDataBoundPos = InstrB(biData, vDataBounds & CByteString("--"))
						
						nPos = InstrB(nDataBoundPos, biData, CByteString("Content-Disposition"))
						nPos = InstrB(nPos, biData, CByteString("name="))
						nPosBegin = nPos + 6
						nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(34)))
						sInputName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
						nPosFile = InstrB(nDataBoundPos, biData, CByteString("filename="))
						nPosBound = InstrB(nPosEnd, biData, vDataBounds)
						
						If nPosFile <> 0 And  nPosFile < nPosBound Then
							Dim oUploadFile, sFileName
							Set oUploadFile = New UploadedFile
							
							nPosBegin = nPosFile + 10
							nPosEnd =  InstrB(nPosBegin, biData, CByteString(Chr(34)))
							sFileName = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
							oUploadFile.FileName = Right(sFileName, Len(sFileName)-InStrRev(sFileName, "\"))

							nPos = InstrB(nPosEnd, biData, CByteString("Content-Type:"))
							nPosBegin = nPos + 14
							nPosEnd = InstrB(nPosBegin, biData, CByteString(Chr(13)))
							
							oUploadFile.ContentType = CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
							
							nPosBegin = nPosEnd+4
							nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
							oUploadFile.FileData = MidB(biData, nPosBegin, nPosEnd-nPosBegin)
							
							If oUploadFile.FileSize > 0 Then Files.Add LCase(sInputName), oUploadFile
						Else
							nPos = InstrB(nPos, biData, CByteString(Chr(13)))
							nPosBegin = nPos + 4
							nPosEnd = InstrB(nPosBegin, biData, vDataBounds) - 2
							If Not mcolFormElem.Exists(LCase(sInputName)) Then mcolFormElem.Add LCase(sInputName), CWideString(MidB(biData, nPosBegin, nPosEnd-nPosBegin))
						End If

						nDataBoundPos = InstrB(nDataBoundPos + LenB(vDataBounds), biData, vDataBounds)
					Loop
				End Sub

				'String to byte string conversion
				Private Function CByteString(sString)
					Dim nIndex
					For nIndex = 1 to Len(sString)
					   CByteString = CByteString & ChrB(AscB(Mid(sString,nIndex,1)))
					Next
				End Function

				'Byte string to string conversion
				Private Function CWideString(bsString)
					Dim nIndex
					CWideString =""
					For nIndex = 1 to LenB(bsString)
					   CWideString = CWideString & Chr(AscB(MidB(bsString,nIndex,1))) 
					Next
				End Function
			End Class

			Class UploadedFile
				Public ContentType
				Public FileName
				Public FileData
				
				Public Property Get FileSize()
					FileSize = LenB(FileData)
				End Property

				Public Sub SaveToDisk(sPath)
					Dim oFS, oFile
					Dim nIndex
				
					If sPath = "" Or FileName = "" Then Exit Sub
					If Mid(sPath, Len(sPath)) <> "\" Then sPath = sPath & "\"
				
					Set oFS = Server.CreateObject("Scripting.FileSystemObject")
					If Not oFS.FolderExists(sPath) Then Exit Sub
					
					Set oFile = oFS.CreateTextFile(sPath & FileName, True)
					
					For nIndex = 1 to LenB(FileData)
					    oFile.Write Chr(AscB(MidB(FileData,nIndex,1)))
					Next

					oFile.Close
				End Sub
				
				Public Sub SaveToDatabase(ByRef oField)
					If LenB(FileData) = 0 Then Exit Sub
					
					If IsObject(oField) Then
						oField.AppendChunk FileData
					End If
				End Sub

			End Class

%>
<%
if session("codusuario")<>"" then
	conectar
	if permisofacultad("admexcluido.asp") then
			'###############################################
			'This shows how to access all files in the html element :
			'Set Uploader = New FileUploader
			'Uploader.Upload()
			'For Each File In FileUploader.Files.Items
			'    Response.Write "File Name:" & File.FileName
			'    Response.Write "File Size:" & File.FileSize
			'    Response.Write "File Type:" & File.ContentType
			'Next

			'This shows how to access file information about a specific file in the html element :
			'Response.Write "File Name:" & FileUploader.Files("file1").FileName
			'Response.Write "File Size:" & FileUploader.Files("file1").FileSize
			'Response.Write "File Type:" & FileUploader.Files("file1").ContentType
			'###############################################


			'------------------------------------------------
			%>
			<!--cargando--><BR><center><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading"><BR><BR><font size=2 color=#483d8b face=Arial>Cargando archivo</font></center><%Response.Flush()%>
			<%
			Dim strFolder, bolUpload, strMessage, strMessage1
			Dim httpref, lngFileSize
			Dim strIncludes, strExcludes

				'-----------------------------------------------
				'name of folder (note there is no / at end)
				strFolder = "c:\files\web\jobs\attachments"
				strFolder = server.MapPath("importados") ''traer desde parametro
				
				'name of folder in http format (note there is no / at end)
				httpRef = "http://localhost/jobs/attachments"
				httpRef = "http://localhost/cobranzacm/importados"  
				'the max size of file which can be uploaded, 0 will give unlimited file size
				lngFileSize = 2147483647  ''10485760 ''(10MB)
				'the files to be excluded (must be in format ".aaa;.bbb")
				'and must be set to blank ("") if none are to be excluded
				strExcludes = ""
				'the files to be included (must be in format ".aaa;.bbb")
				'and must be set to blank ("") if none are to be excluded
				''strIncludes = ".gif;.jpg;.jpeg;.txt;.pdf;.doc;.avi;.mpeg;.mpg;.mp3;.wma;.xls;.ppt;.bmp;.png;.tiff;"
				strIncludes = ".txt;.TXT;"
				'-----------------------------------------------

			' Create the FileUploader
			Dim Uploader, File
			Set Uploader = New FileUploader

			' This starts the upload process
			Uploader.Upload()

			'******************************************
			' Use [FileUploader object].Form to access 
			' additional form variables submitted with
			' the file upload(s). (used below)
			'******************************************
				' Check if any files were uploaded
				Dim cuentaarch
				If Uploader.Files.Count = 0 Then
					strMessage = "No file entered."
					strMessage1= "No file entered."
				Else
					cuentaarch=0
					' Loop through the uploaded files
					For Each File In Uploader.Files.Items		

						bolUpload = false		

						'Response.Write lngMaxSize
						'Response.End 

						if lngFileSize = 0 then
							bolUpload = true
						else		
							if File.FileSize > lngFileSize then
								bolUpload = false
								strMessage = "File too large"
							else
								bolUpload = true
							end if
						end if

						if bolUpload = true then				
						    'Check to see if file extensions are excluded
						    If strExcludes <> "" Then
								If ValidFileExtension(File.FileName, strExcludes) Then
						            strMessage = "It is not allowed to upload a file containing a [." & GetFileExtension(File.FileName) & "] extension"
									bolUpload = false
								End If
							End If
							'Check to see if file extensions are included
							If strIncludes <> "" Then
								If InValidFileExtension(File.FileName, strIncludes) Then
									strMessage = "It is not allowed to upload a file containing a [." & GetFileExtension(File.FileName) & "] extension"
									bolUpload = false
								End If
							End If			
						end if
						
					

						if bolUpload = true then
							if cuentaarch=0 then
								sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUpload' or descripcion='RutaWebUpload'"
								consultar sql,RS
								RS.Filter=" descripcion='RutaFisicaUpload'"
								RutaFisicaUpload=RS.Fields(1)
								RS.Filter=" descripcion='RutaWebUpload'"
								RutaWebUpload=RS.Fields(1)				
								RS.Filter=""
								RS.Close	
									
									strFolder=RutaFisicaUpload	
									
									
								sql="select isnull(correo,'') as correo from usuario where codusuario=" & session("codusuario")
								consultar sql,RS
								correo=RS.Fields("correo")
								RS.Close
									
								if Uploader.Form("activo")<>"" then ''Estado Pendiente actualización total
									sql="INSERT INTO CargaArchivo(proceso,rutaarchivo,mailinforme,usuarioregistra,fecharegistra,codestadocarga) VALUES ('Excluidos','" & RutaFisicaUpload & "\excluidos','" & correo & "'," & session("codusuario") & ",getdate(),2)"
								else ''Estado Pendiente actualización inserción
									sql="INSERT INTO CargaArchivo(proceso,rutaarchivo,mailinforme,usuarioregistra,fecharegistra,codestadocarga) VALUES ('Excluidos','" & RutaFisicaUpload & "\excluidos','" & correo & "'," & session("codusuario") & ",getdate(),1)"
								end if
								conn.execute sql		
								
								''capturamos el maximo CargaArchivo para el FileName que sera "excluidos(nro).txt"
								sql="select max(codcarga) as codcarga from CargaArchivo where usuarioregistra=" & session("codusuario")
								consultar sql,RS
								codcarga=RS.Fields("codcarga")
								RS.Close
								
								''agregamos el nro a la ruta del registro creado
								sql="update CargaArchivo set rutaarchivo=rutaarchivo + '" & codcarga & ".txt' where codcarga=" & codcarga
								conn.execute sql
								
								''Esto levanta el archivo a la carpeta
								File.FileName="excluidos" & codcarga & "." & GetFileExtension(File.FileName)
								''File.SaveToDisk strFolder ' Save the file
								File.SaveToDisk Server.MapPath("temporal\")			
								''strMessage =  "File Uploaded: " & File.FileName
								''strMessage = File.FileName
								strMessage = "OK"
								'strMessage = strMessage & "Size: " & File.FileSize & " bytes<br>"
								'strMessage = strMessage & "Type: " & File.ContentType & "<br><br>"										
											
							end if
						end if
					cuentaarch=cuentaarch + 1
					Next
					
					
					'
					Dim name
					
				    name = 	Uploader.Form("txtName")    'Used to extract fields in the form
					
				End If
			'Response.Redirect ("uploadform.asp?msg=" & strMessage)
			''Response.Redirect "cargarexcluido.asp?agregardato=" & Uploader.Form("agregardato") & "&flaginactivar=" & Uploader.Form("flaginactivar") & "&archivotxt=" & strMessage
			%>
				<script language=javascript>
					window.open("cargarexcluido.asp?agregardato=<%=Uploader.Form("agregardato")%>&flaginactivar=<%=Uploader.Form("flaginactivar")%>&archivotxt=<%=strMessage%>","_self");
				</script>
			<%

			''Response.Write strMessage 
			''Response.Write Uploader.Form("idotoscopia") 

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

<%


			'--------------------------------------------
			' ValidFileExtension()
			' You give a list of file extensions that are allowed to be uploaded.
			' Purpose:  Checks if the file extension is allowed
			' Inputs:   strFileName -- the filename
			'           strFileExtension -- the fileextensions not allowed
			' Returns:  boolean
			' Gives False if the file extension is NOT allowed
			'--------------------------------------------
			Function ValidFileExtension(strFileName, strFileExtensions)

			    Dim arrExtension
			    Dim strFileExtension
			    Dim i
			    
			    strFileExtension = UCase(GetFileExtension(strFileName))
			    
			    arrExtension = Split(UCase(strFileExtensions), ";")
			    
			    For i = 0 To UBound(arrExtension)
			        
			        'Check to see if a "dot" exists
			        If Left(arrExtension(i), 1) = "." Then
			            arrExtension(i) = Replace(arrExtension(i), ".", vbNullString)
			        End If
			        
			        'Check to see if FileExtension is allowed
			        If arrExtension(i) = strFileExtension Then
			            ValidFileExtension = True
			            Exit Function
			        End If
			        
			    Next
			    
			    ValidFileExtension = False

			End Function

			'--------------------------------------------
			' InValidFileExtension()
			' You give a list of file extensions that are not allowed.
			' Purpose:  Checks if the file extension is not allowed
			' Inputs:   strFileName -- the filename
			'           strFileExtension -- the fileextensions that are allowed
			' Returns:  boolean
			' Gives False if the file extension is NOT allowed
			'--------------------------------------------
			Function InValidFileExtension(strFileName, strFileExtensions)

			    Dim arrExtension
			    Dim strFileExtension
			    Dim i
			        
			    strFileExtension = UCase(GetFileExtension(strFileName))
			    
			    'Response.Write "filename : " & strFileName & "<br>"
			    'Response.Write "file extension : " & strFileExtension & "<br>"    
			    'Response.Write strFileExtensions & "<br>"
			    'Response.End 
			    
			    arrExtension = Split(UCase(strFileExtensions), ";")
			    
			    For i = 0 To UBound(arrExtension)
			        
			        'Check to see if a "dot" exists
			        If Left(arrExtension(i), 1) = "." Then
			            arrExtension(i) = Replace(arrExtension(i), ".", vbNullString)
			        End If
			        
			        'Check to see if FileExtension is not allowed
			        If arrExtension(i) = strFileExtension Then
			            InValidFileExtension = False
			            Exit Function
			        End If
			        
			    Next
			    
			    InValidFileExtension = True

			End Function

			'--------------------------------------------
			' GetFileExtension()
			' Purpose:  Returns the extension of a filename
			' Inputs:   strFileName     -- string containing the filename
			'           varContent      -- variant containing the filedata
			' Outputs:  a string containing the fileextension
			'--------------------------------------------
			Function GetFileExtension(strFileName)

			    GetFileExtension = Mid(strFileName, InStrRev(strFileName, ".") + 1)
			    
			End Function

%>



