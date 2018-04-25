<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<%
if session("codusuario")<>"" then
		conectar
		datapersona = request("datapersona")
		IDCampana = request("IDCampana")
		telefonoactivo = request("telefonoactivo")
		idCampPerTelf = request("idCampPerTelf")
		
		idcamperacc = request("idcamperacc") 
		idgestion = request("idgestion")
		comentario = request("comentario")
		tpress = request("tpress")
		idaccionactiva = request("idaccionactiva")	
		codtipocontacto = request("codtipocontacto")	
		
		DTELEFONO = telefonoactivo
		
		sqlTransMysql = "select * from OPENQUERY(G2, 'SELECT vlog.uniqueid, " & chr(10) & _
						"vlog.lead_id, " & chr(10) & _
						"vlog.list_id, " & chr(10) & _
						"vlog.campaign_id, " & chr(10) & _
						"vlog.call_date, " & chr(10) & _
						"vlog.start_epoch, " & chr(10) & _
						"vlog.end_epoch, " & chr(10) & _
						"vlog.length_in_sec, " & chr(10) & _
						"vlog.phone_number, " & chr(10) & _
						"vlog.user, " & chr(10) & _
						"vlog.called_count, " & chr(10) & _
						"rlog.filename, " & chr(10) & _
						"rlog.server_ip " & chr(10) & _ 						
					"FROM vicidial_log vlog " & chr(10) & _
					"INNER JOIN recording_log rlog ON vlog.lead_id=rlog.lead_id " & chr(10) & _
					"WHERE vlog.phone_number = '944267001' " & chr(10) & _
					"AND vlog.campaign_id = '18021801' " & chr(10) & _
					"AND vlog.user=rlog.user " & chr(10) & _
					"GROUP BY vlog.call_date " & chr(10) & _
					"ORDER BY vlog.call_date DESC  " & chr(10) & _
					"LIMIT 1')"
					
			'response.write sqlTransMysql
					
			consultar sqlTransMysql,RS6

			uniqueid      = RS6.fields("uniqueid")
			lead_id       = RS6.fields("lead_id")
			list_id       = RS6.fields("list_id")
			campaign_id   = RS6.fields("campaign_id")
			call_date     = RS6.fields("call_date")
			start_epoch   = RS6.fields("start_epoch")
			end_epoch     = RS6.fields("end_epoch")
			length_in_sec = RS6.fields("length_in_sec")
			phone_number  = RS6.fields("phone_number")
			uuser         = RS6.fields("user")
			called_count  = RS6.fields("called_count")
			filename      = RS6.fields("filename")
			server_ip     = RS6.fields("server_ip")
			
			
			RS6.Close
		
		

		if idcamperacc = "" then

					sql = "INSERT INTO [dbo].[Campaña_Persona_Accion] " & chr(10) & _
					           "([FechaHoraInicioGestion] " & chr(10) & _
					           ",[FechaHoraFinGestion]" & chr(10) & _
					           ",[FechaRegistra]" & chr(10) & _
					           ",[UsuarioRegistra]" & chr(10) & _
					           ",[FechaModifica]" & chr(10) & _
					           ",[UsuarioModifica]" & chr(10) & _
					           ",[IDCampañaPersona]" & chr(10) & _
					           ",[IDTipoAccion]" & chr(10) & _
					           ",[Estado]" & chr(10) & _
					           ",[Alerta]" & chr(10) & _
					           ",[SpeechPersonalizado]" & chr(10) & _
					           ",[AudioPersonalizado]" & chr(10) & _					           
					           ",[IDCampañaPersonaTelefono]" & chr(10) & _
					           ",[IDCampañaPersonaDireccion]" & chr(10) & _
					           ",[IDCampañaPersonaEmail]" & chr(10) & _
					           ",[CostoSoles]" & chr(10) & _
					           ",[UsuarioAsignado]" & chr(10) & _
					           ",[UsuarioEjecutor]" & chr(10) & _
					           ",[DuracionCall]" & chr(10) & _
					           ",[TipoDial]" & chr(10) & _
					           ",[Comentario]" & chr(10) & _
					           ",[IDGestion]" & chr(10) & _
					           ",[IDMotivo]" & chr(10) & _
					           ",[IDTipoContacto])" & chr(10) & _
					     "VALUES" & chr(10) & _
					      "(GETDATE() " & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",GETDATE()" & chr(10) & _
					           ","& session("codusuario") & " "& chr(10) & _
					           ",NULL" & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",'" & datapersona & "'" & chr(10) & _
					           ",1" & chr(10) & _
					           ",1" & chr(10) & _
					           ",0" & chr(10) & _					          
					           ",NULL" & chr(10) & _
					           ",NULL" & chr(10) & _
					           "," & idCampPerTelf & "" & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",NULL" & chr(10) & _           
					           ",(SELECT UsuarioAsignado FROM Campaña_Persona where IDCampañaPersona = " & datapersona & ")" & chr(10) & _
					           "," & session("codusuario") & " " & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",1" & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",NULL" & chr(10) & _
					           ",3)"				          
					         

					        conn.Execute sql

						  sql=   "Select MAX(IDCampañaPersonaAccion) as Id from Campaña_Persona_Accion where UsuarioRegistra = " & session("codusuario")

						  
						  consultar sql,RS6

						  res = RS6.fields("Id")
						
						 RS6.Close
						 
				sql="update Campaña_Persona_Accion" & chr(10) & _
								"set goUniqueid="& uniqueid & chr(10) & _
								",goLead_id="& lead_id & chr(10) & _
								",goList_id="& list_id & chr(10) & _
								",goCampaign_id="& campaign_id & chr(10) & _
								",goCall_date="& call_date & chr(10) & _
								",goStart_epoch="& start_epoch & chr(10) & _
								",goEnd_epoch="& end_epoch & chr(10) & _
								",goLength_in_sec="& length_in_sec & chr(10) & _
								",goPhone_number="& phone_number & chr(10) & _
								",goUser="& uuser & chr(10) & _
								",goCalled_count="& called_count & chr(10) & _
								",goFilename="& filename & chr(10) & _
								",goServer_ip="& server_ip & chr(10) & _									
								"where IDCampañaPersonaAccion = " & res

								'response.write sql

					        conn.Execute sql
						 


					%>
					<table class="tabinterna"  id="tabinterna_gestion" valign="top">
														<tr class="cabecera-orange" valign="top">
															<td  colspan="2" >
																<input type="hidden" name="idgestion" id="idgestion" value="<%=res%>" >
																Agregar Gesti&oacute;n <%=res%>
															</td>										
														</tr> 
														<tr class="fondo-blanco">
															<td class="text-orange">
															Acci&oacute;n Activa
															</td>
															<td class="text-orange">
																<p id="text-accion"><%=tpress%></p>
															</td>	
														</tr>	
														<tr class="fondo-blanco">
															<td class="text-orange">
															Telefono
															</td>
															<td class="text-orange">
																<%=telefonoactivo%>
															</td>	
														</tr>
														<tr class="fondo-blanco">
															<td class="text-orange">
															Tipo Contacto
															</td>
															<td class="text-orange">
																<select style="font-size: 11.5px;" id="codtipocontacto" name="codtipocontacto">
																	<option value="">Seleccione un Tipo contacto</option>
																	<%if datapersona <> "" then
																	sql = " SELECT IDTipoContacto, Descripcion FROM TipoContacto "
																	consultar sql,RS4
																	DO while not RS4.EOF
																	 %>
																	<option value="<%=RS4.fields("IDTipoContacto")%>"><%=RS4.fields("Descripcion")%></option>
																	<% 
																	RS4.MoveNext
																	Loop
																	RS4.Close

																	end if%>
																</select>
															</td>	
														</tr>	
														<tr class="fondo-blanco">
															<td class="text-orange">
															Respuesta
															</td>
															<td class="text-orange">
																<select style="font-size: 11.5px;" id="codigogestion">
																	<option value="">Seleccione una Respuesta</option>
																	<%if datapersona <> "" then
																	sql = " SELECT a.IDGestion, a.Descripcion FROM Gestion a where IDTipoCampaña = (Select IDTipoCampaña FROM Campaña WHERE idcampaña =" & IDCampana & ")"

																	response.write sql

																	consultar sql,RS4
																	DO while not RS4.EOF
																	%>
																	<option value="<%=RS4.fields("IDGestion")%>"><%=RS4.fields("Descripcion")%></option>
																	<% 
																	RS4.MoveNext
																	Loop
																	RS4.Close
																	end if%>
																</select>
															</td>	
														</tr>
														<tr class="fondo-red">
															<td class="text-withe" colspan="2">
															Comentario
															</td>											
														</tr>		
														<tr>
														<td class="text-orange" colspan="2">
																<textarea class="areatexto" name="comentario" id="comentario"></textarea>
														</td>	
														</tr>
														<tr class="fondo-red">
															<td style="text-align: right; width: 100%;" colspan="2"><a href="#" onclick = "javascript:guardargestion('<%=res%>','<%=datapersona%>','<%=idaccionactiva%>')"><i class="demo-icon icon-floppy">&#xe809;</i><a></td>
														<tr>				
													</table>
					<%
					desconectar

					else


							sql="update Campaña_Persona_Accion" & chr(10) & _
								"set FechaHoraFinGestion = getdate() ," & chr(10) & _
									"FechaModifica = getdate()," & chr(10) & _
									"UsuarioModifica = " & session("codusuario") & " ," & chr(10) & _
									"comentario = '" & comentario & "'," & chr(10) & _
									"IDGestion = " & idgestion & "," & chr(10) & _	
									"IDTipoContacto = " & codtipocontacto & "" & chr(10) & _									
								"where IDCampañaPersonaAccion = " & idcamperacc

								'response.write sql

					        conn.Execute sql

							sqlU="update Campaña_Persona_Accion" & chr(10) & _
								"set goUniqueid="& uniqueid & chr(10) & _
								",goLead_id="& lead_id & chr(10) & _
								",goList_id="& list_id & chr(10) & _
								",goCampaign_id="& campaign_id & chr(10) & _
								",goCall_date="& call_date & chr(10) & _
								",goStart_epoch="& start_epoch & chr(10) & _
								",goEnd_epoch="& end_epoch & chr(10) & _
								",goLength_in_sec="& length_in_sec & chr(10) & _
								",goPhone_number="& phone_number & chr(10) & _
								",goUser="& uuser & chr(10) & _
								",goCalled_count="& called_count & chr(10) & _
								",goFilename="& filename & chr(10) & _
								",goServer_ip="& server_ip & chr(10) & _									
								"where IDCampañaPersonaAccion = " & idcamperacc

								'response.write sql

					        conn.Execute sqlU
					        	

									  


					%>
					<table class="tabinterna"  id="tabinterna_gestion" valign="top">
														<tr class="cabecera-orange" valign="top">
															<td  colspan="2" >
																<input type="hidden" name="idgestion" id="idgestion" value="" >
																Agregar Gesti&oacute;n
															</td>										
														</tr> 
														<tr class="fondo-blanco">
															<td class="text-orange">
															Acci&oacute;n Activa
															</td>
															<td class="text-orange">
																<p id="text-accion">En espera</p>
															</td>	
														</tr>	
														<tr class="fondo-blanco">
															<td class="text-orange">
															Telefono
															</td>
															<td class="text-orange">																
															</td>	
														</tr>
														<tr class="fondo-blanco">
															<td class="text-orange">
															Tipo Contacto
															</td>
															<td class="text-orange">
																<select style="font-size: 11.5px;" id="codtipocontacto" name="codtipocontacto">
																	<option value="">Seleccione un Tipo contacto</option>
																	<%if datapersona <> "" then
																	sql = " SELECT IDTipoContacto, Descripcion FROM TipoContacto "
																	consultar sql,RS4
																	DO while not RS4.EOF
																	 %>
																	<option value="<%=RS4.fields("IDTipoContacto")%>"><%=RS4.fields("Descripcion")%></option>
																	<% 
																	RS4.MoveNext
																	Loop
																	RS4.Close

																	end if%>
																</select>
															</td>	
														</tr>	
														<tr class="fondo-blanco">
															<td class="text-orange">
															Respuesta
															</td>
															<td class="text-orange">
																<select style="font-size: 11.5px;">
																	<option value="">Seleccione una Respuesta</option>
																	<%if datapersona <> "" then
																	sql = " SELECT IDGestion, Descripcion FROM Gestion where IDTipoCampaña = (Select IDTipoCampaña FROM Campaña WHERE idcampaña =" & IDCampana & ")"
																	consultar sql,RS4
																	DO while not RS4.EOF
																	%>
																	<option value<%=RS4.fields("IDGestion")%>><%=RS4.fields("Descripcion")%></option>
																	<% 
																	RS4.MoveNext
																	Loop
																	RS4.Close
																	end if%>
																</select>
															</td>	
														</tr>
														<tr class="fondo-red">
															<td class="text-withe" colspan="2">
															Comentario
															</td>											
														</tr>		
														<tr>
														<td class="text-orange" colspan="2">
																<textarea class="areatexto" name="comentario"></textarea>
														</td>	
														</tr>
														<tr class="fondo-red">
															<td style="text-align: right; width: 100%;" colspan="2"><a href="#" ><i class="demo-icon icon-floppy">&#xe809;</i><a></td>
														<tr>				
													</table>
					<%
					desconectar



					end if

end if
%>