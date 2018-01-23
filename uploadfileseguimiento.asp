<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<%
			'***************************************
			' File:	  Upload.asp
			' Author: Jacob "Beezle" Gilley
			' direccion:  avis7@airmail.net
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
	if permisofacultad("admimpagado.asp") then
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
			<!--cargando--><BR><center><img src=imagenes/loading.gif border=0 id="imgloading" name="imgloading"><BR><BR><font size=2 color=#483d8b face=Arial>Cargando archivo</font></center><%''Response.Flush()%>
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
				lngFileSize = 2147483647 ''(10MB) /2147483647  ''
				'the files to be excluded (must be in format ".aaa;.bbb")
				'and must be set to blank ("") if none are to be excluded
				strExcludes = ""
				'the files to be included (must be in format ".aaa;.bbb")
				'and must be set to blank ("") if none are to be excluded
				strIncludes = ".gif;.jpg;.jpeg;.txt;.pdf;.doc;.docx;.avi;.mpeg;.mpg;.mp3;.mp4;.wma;.wav;.xls;.xlsx;.ppt;.pptx;.bmp;.png;.tiff;.zip;.rar;"
                ''strIncludes = ".txt;.TXT;"
                ''strIncludes = ""
				'-----------------------------------------------

			' Create the FileUploader
			Dim Uploader, File
			Set Uploader = New FileUploader

			' This starts the upload process
			Uploader.Upload()

			codtipoaccion=Uploader.Form("codtipoaccion_data")
			codtipocontacto=Uploader.Form("codtipocontacto_data")
			coddirtel=Uploader.Form("coddirtel_data")
			codgestion=Uploader.Form("codgestion_data")
			comentario=Uploader.Form("comentario_data")
			fechapromesa=Uploader.Form("fechapromesa_data")
			divisa1=Uploader.Form("divisa1_data")
			importe1=Uploader.Form("importe1_data")
			divisa2=Uploader.Form("divisa2_data")
			importe2=Uploader.Form("importe2_data")				
			vistapadre1=Uploader.Form("vistapadre1_data")
			paginapadre1=Uploader.Form("paginapadre1_data")
			vistapadre=Uploader.Form("vistapadre_data")
			paginapadre=Uploader.Form("paginapadre_data")
			codigocentral=Uploader.Form("codigocentral_data")
			contrato=Uploader.Form("contrato_data")
			fechadatos=Uploader.Form("fechadatos_data")
			fechagestion=Uploader.Form("fechagestion_data")		
			
			sql="select count(*) from Gestion where codgestion='" & codgestion & "' and UPPER(descripcion) like '%PROMESA%' and activo=1"
			consultar sql,RS
			espromesa=RS.Fields(0)
			RS.Close	
			
			if codtipoaccion<>"2" then ''no visita

				if replace(coddirtel,"fono","")<>coddirtel then
				
						''sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
						''consultar sql,RS	
						''maxfechagestion=rs.fields(0)
						''RS.Close	
		
						''if mid(Uploader.Form("fechagestion"),7,4) & mid(Uploader.Form("fechagestion"),4,2) & mid(Uploader.Form("fechagestion"),1,2)=CStr(maxfechagestion) then
						''	vistabusqueda="VERULTIMPAGADO"
						''else
						''	vistabusqueda="VERIMPAGADO"
						''end if		
				
				        vistabusqueda="ClienteDiario"
				        
						sql="select top 1 tipofono1,prefijo1,fono1,extension1,tipofono2,prefijo2,fono2,extension2,tipofono3,prefijo3,fono3,extension3,tipofono4,prefijo4,fono4,extension4,tipofono5,prefijo5,fono5,extension5 from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "'"
						''Response.Write sql
						consultar sql,RS1	
						if not RS1.EOF then
							if coddirtel="fono1" then
								tipofono=RS1.Fields("tipofono1")
								prefijo=RS1.Fields("prefijo1")
								fono=RS1.Fields("fono1")
								extension=RS1.Fields("extension1")
							end if
							if coddirtel="fono2" then
								tipofono=RS1.Fields("tipofono2")
								prefijo=RS1.Fields("prefijo2")
								fono=RS1.Fields("fono2")
								extension=RS1.Fields("extension2")
							end if
							if coddirtel="fono3" then
								tipofono=RS1.Fields("tipofono3")
								prefijo=RS1.Fields("prefijo3")
								fono=RS1.Fields("fono3")
								extension=RS1.Fields("extension3")
							end if
							if coddirtel="fono4" then
								tipofono=RS1.Fields("tipofono4")
								prefijo=RS1.Fields("prefijo4")
								fono=RS1.Fields("fono4")
								extension=RS1.Fields("extension4")
							end if
							if coddirtel="fono5" then
								tipofono=RS1.Fields("tipofono5")
								prefijo=RS1.Fields("prefijo5")
								fono=RS1.Fields("fono5")
								extension=RS1.Fields("extension5")
							end if
						end if		
						RS1.Close				
				else
						sql="select * from TelefonoNuevo A where A.codigocentral='" & codigocentral & "' and codtelefononuevo=" & coddirtel
						''Response.Write sql
						consultar sql,RS1	
						if not RS1.EOF then
							tipofono=RS1.Fields("codtipotelefono")
							prefijo=RS1.Fields("prefijo")
							fono=RS1.Fields("fono")
							extension=RS1.Fields("extension")
						end if		
						RS1.Close						
				end if
				
			else ''visita
			
			
				if coddirtel="dirprin" then
				
						''sql="select max(convert(varchar,fechagestion,112)) from DistribucionDiaria"
						''consultar sql,RS	
						''maxfechagestion=rs.fields(0)
						''RS.Close	
		
						''if mid(Uploader.Form("fechagestion"),7,4) & mid(Uploader.Form("fechagestion"),4,2) & mid(Uploader.Form("fechagestion"),1,2)=CStr(maxfechagestion) then
						''	vistabusqueda="VERULTIMPAGADO"
						''else
						''	vistabusqueda="VERIMPAGADO"
						''end if		
						
						vistabusqueda="ClienteDiario"
				
						sql="select top 1 direccion,codestado as departamento,codprovincia as provincia,coddistrito as distrito from " & vistabusqueda & " A where A.codigocentral='" & codigocentral & "' and fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "'"
						''Response.Write sql
						consultar sql,RS1	
						if not RS1.EOF then
							direccion=RS1.Fields("direccion")
							coddpto=RS1.Fields("departamento")
							codprov=RS1.Fields("provincia")
							coddist=RS1.Fields("distrito")
						end if		
						RS1.Close		

				else
						sql="select * from DireccionNueva A where A.codigocentral='" & codigocentral & "' and coddireccionnueva=" & coddirtel
						''Response.Write sql
						consultar sql,RS1	
						if not RS1.EOF then
							direccion=RS1.Fields("direccion")
							coddpto=RS1.Fields("coddpto")
							codprov=RS1.Fields("codprov")
							coddist=RS1.Fields("coddist")
						end if		
						RS1.Close						
				end if
							
			
			end if
			
			sql="select codagencia from usuario where codusuario=" & session("codusuario")
			consultar sql,RS
			if IsNull(RS.Fields(0)) then
				codagencia="null"
			else
				codagencia=CStr(RS.Fields(0))
			end if
			RS.Close				
			
			''La Gestión es por cliente así que aquí se tiene que recorrer todos sus contratos distribuidos a esa agencia
			sql="select contrato from DistribucionDiaria where codigocentral='" & codigocentral & "' and fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "'"
			consultar sql,RS1
			Do While Not RS1.EOF
			    ''Response.Write "<BR>inicia insercion " & now() & "<BR>"
			    if espromesa>0 then
				    sql="insert into RespuestaGestion (codigocentral,fechadatos,contrato,fhgestionado,codagencia,codgestion,comentario,usuarioregistra,fecharegistra,tipofono,prefijo,fono,extension,direccion,coddpto,codprov,coddist,fechapromesa,divisa1,importe1,divisa2,importe2) " & _
					    "values ('" & codigocentral & "','" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "','" & RS1.Fields("contrato") & "',getdate()," & codagencia & ",'" & codgestion & "','" & comentario & "'," & session("codusuario") & ",getdate(),'" & tipofono & "','" & prefijo & "','" & fono & "','" & extension & "','" & direccion & "','" & coddpto & "','" & codprov & "','" & coddist & "','" & mid(fechapromesa,7,4) & mid(fechapromesa,4,2) & mid(fechapromesa,1,2) & "','" & divisa1 & "'," & iif(importe1="","null",replace(importe1,",",""))  & ",'" & divisa2 & "'," & iif(importe2="","null",replace(importe2,",",""))  & ") "
			    else
				    sql="insert into RespuestaGestion (codigocentral,fechadatos,contrato,fhgestionado,codagencia,codgestion,comentario,usuarioregistra,fecharegistra,tipofono,prefijo,fono,extension,direccion,coddpto,codprov,coddist) " & _
					    "values ('" & codigocentral & "','" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "','" & RS1.Fields("contrato") & "',getdate()," & codagencia & ",'" & codgestion & "','" & comentario & "'," & session("codusuario") & ",getdate(),'" & tipofono & "','" & prefijo & "','" & fono & "','" & extension & "','" & direccion & "','" & coddpto & "','" & codprov & "','" & coddist & "') "
			    end if
			    ''Response.Write sql
			    conn.execute sql
                ''Response.Write "<BR>inserto " & now() & "<BR>"
			    
			    sql="select top 1 A.codrespgestion from RespuestaGestion A inner join Gestion B ON A.codgestion = B.codgestion inner join TipoContacto C ON B.codtipocontacto = C.codtipocontacto where A.codigocentral='" & codigocentral & "' and A.contrato='" & RS1.Fields("contrato") & "' and A.fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "' ORDER BY C.codtipocontacto, isNull(A.fechapromesa, '99991231'), A.fhgestionado DESC"
			    ''response.Write sql
			    consultar sql,RS2
				mejor_codrespgestion=RS2.Fields(0)
			    RS2.Close		
			    
			    ''Response.Write "<BR>busco mejor gestion " & now() & "<BR>"			    
			    

			    ''Actualizamos la mejor gestion
                sql="update ContratoDiario set mejor_codrespgestion=" & mejor_codrespgestion & " where codigocentral='" & codigocentral & "' and contrato='" & RS1.Fields("contrato") & "' and fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "'"
                conn.execute sql

                sql="update UltContratoDiario set mejor_codrespgestion=" & mejor_codrespgestion & " where codigocentral='" & codigocentral & "' and contrato='" & RS1.Fields("contrato") & "' and fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "'"
                conn.execute sql
                
				sql="update BusquedaDistribucion set mejor_codrespgestion=" & mejor_codrespgestion & " where codigocentral='" & codigocentral & "' and contrato='" & RS1.Fields("contrato") & "' and fechadatos='" & mid(fechadatos,7,4) & mid(fechadatos,4,2) & mid(fechadatos,1,2) & "'"
				conn.execute sql

			    ''Response.Write "<BR>actualizo mejor gestion " & now() & "<BR>"			    
			
			        ''El fichero se sube a la gestion del ultimo contrato registrado
			        sql="select top 1 codrespgestion from RespuestaGestion where usuarioregistra=" & session("codusuario") & " order by fecharegistra desc"
			        sql="select max(codrespgestion) from RespuestaGestion where usuarioregistra=" & session("codusuario")
			        consultar sql,RS
			        codrespgestion=RS.Fields(0)
			        RS.Close				
                ''Response.Write "<BR>busco gestion insertada " & now() & "<BR>"		
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
								        sql="select descripcion,valortexto1 from parametro where descripcion='RutaFisicaUploadExt' or descripcion='RutaWebUpload'"
								        consultar sql,RS
								        RS.Filter=" descripcion='RutaFisicaUploadExt'"
								        RutaFisicaUpload=RS.Fields(1)
								        RS.Filter=" descripcion='RutaWebUpload'"
								        RutaWebUpload=RS.Fields(1)				
								        RS.Filter=""
								        RS.Close	
        								
								        strFolder=RutaFisicaUpload
        									
								        ''Esto levanta el archivo a la carpeta
								        File.FileName="adjgestion" & codrespgestion & "." & GetFileExtension(File.FileName)
								        ''File.SaveToDisk "c:\archivos\" ' Save the file			
								        ''File.SaveToDisk "\\118.180.22.116\xcobranzas\Extranet\ImportadosWeb\"
								        ''File.SaveToDisk "W:\Extranet\ImportadosWeb\"

								        File.SaveToDisk Server.MapPath("temporal\")
        								
								        ''rutabat=(Server.MapPath("moveradjuntos.bat")
								        ''Shell (rutabat)
							            ''set wshell = CreateObject("WScript.Shell") 
                                        ''wshell.run "C:\inetpub\wwwroot\cobranzacm\moveradjuntos.bat" ''Server.MapPath("cobranzacm\moveradjuntos.bat")
                                        ''set wshell = nothing 
								        ''set ObjBat = server.createobject("WScript.shell")
                                        ''ObjBat.Run(Server.MapPath("cobranzacm\moveradjuntos.bat"), 2, true)
                                        
                                        ''dim fs
                                        ''set fs=Server.CreateObject("Scripting.FileSystemObject")
                                        ''fs.MoveFile "c:\archivos\" & File.FileName,"W:\Extranet\ImportadosWeb\" & File.FileName
                                        ''set fs=nothing	
                                        							
                                        ''strMessage =  "File Uploaded: " & File.FileName
								        ''strMessage = File.FileName
								        strMessage = "OK"
								        'strMessage = strMessage & "Size: " & File.FileSize & " bytes<br>"
								        'strMessage = strMessage & "Type: " & File.ContentType & "<br><br>"										
        											
								        sql="Update RespuestaGestion set ficherogestion='" & File.FileName & "' where codrespgestion=" & codrespgestion
								        conn.execute sql
							        end if
						        end if
					        cuentaarch=cuentaarch + 1
					        Next
        					
        					
					        '
					        Dim name
        					
				            name = 	Uploader.Form("txtName")    'Used to extract fields in the form
        					
				        End If
				
			RS1.MoveNext
            Loop
			RS1.Close				
			'Response.Redirect ("uploadform.asp?msg=" & strMessage)
			''Response.Redirect "cargadireccion.asp?agregardato=" & Uploader.Form("agregardato") & "&flaginactivar=" & Uploader.Form("flaginactivar") & "&archivotxt=" & strMessage
			%>
				<script language=javascript>
					window.open("<%=paginapadre1%>?vistapadre=<%=vistapadre%>&paginapadre=<%=paginapadre%>&codigocentral=<%=codigocentral%>&contrato=<%=contrato%>&fechadatos=<%=fechadatos%>&fechagestion=<%=fechagestion%>","<%=vistapadre1%>");
					window.close();
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



