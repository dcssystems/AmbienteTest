<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<%
if session("codusuario")<>"" then
conectar
datapersona = request("datapersona")
%>

<table class="tabinterna"  id="tabinterna_telf">
														<tr class="cabecera-orange">
															<td colspan="6">Telefonos
															</td>
														</tr>
														<tr class="fondo-red">
															<td class="text-withe">Tipo</td>
															<td class="text-withe">Prf</td>
															<td class="text-withe">Número</td>
															<td class="text-withe">Ext</td>
															<td class="text-withe">Descripción</td>
															<td class="text-withe"></td>
														</tr>
														<%
														if datapersona <> "" then

														varcolor = 0
															
														sql = "select a.IDCampañaPersonaTelefono,(select descripcion from TipoTelefono where IDTipoTelefono = a.IDTipoTelefono) as Tipo,a.Prefijo,a.Numero,a.Extension,a.Descripcion from Campaña_Persona_Telefono a where IDCampañaPersona =" & datapersona

														consultar sql,RS4
														Do While Not RS4.EOF	 
														%>
														<tr class="fondo-red <% IF varcolor	 = 0 Then %> fondo-blanco <% Else %> fondo-rojo <% End IF %>" >
															<input type="hidden" name="<%=RS4.Fields("Numero")%>" id="<%=RS4.Fields("Numero")%>" value="<%=RS4.Fields("IDCampañaPersonaTelefono")%>" />
															<td><%=RS4.Fields("Tipo")%></td>
															<td><%=RS4.Fields("Prefijo")%></td>
															<td><%=RS4.Fields("Numero")%></td>
															<td><%=RS4.Fields("Extension")%></td>
															<td><%=RS4.Fields("Descripcion")%></td>
															<td style="background: #a42627; text-align: center;"><a href="#" onclick="javascript:creargestion(<%=datapersona%>,<%=idcampana%>,<%=RS4.Fields("Numero")%>)"><div><i class="demo-icon2  icon-phone-circled" style="color:#FE6D2E !important;" >&#xe822;</i></div></a></td>
														</tr>
														<%
														IF varcolor	 = 0 Then
															varcolor	 = 1
														else
															varcolor	 = 0 
														end if
														RS4.MoveNext
														loop
														RS4.Close
														end if
														%>
														<tr class="fondo-red">
															<td class="text-withe" >
																<select style="font-size: 11.5px;" name="tipotelf" id="tipotelf"><option value="">Tipo Tel</option>
																<% 
																if datapersona <> "" then

																sql = "Select IDTipoTelefono, Descripcion from TipoTelefono"
																consultar sql,RS4

																Do while Not RS4.EOF
																%>
																<option value="<%=RS4.fields("IDTipoTelefono")%>"><%=RS4.fields("Descripcion")%></option>
																<%
																RS4.MoveNext
															    Loop
															    RS4.Close
																end if
																%>
																</select>

															</td>
															<td class="text-withe">51</td>
															<td><input type="text" id="telnuevo" style="width: 72px; font-size: 11.5px;" name="telnuevo" /></td>
															<td>
																<input type="text" id="exttelnuevo" style="width: 50px; font-size: 11.5px;" name="exttelnuevo" />
															</td>
															<td>
																<input type="text" id="destelnuevo" style="width: 100px; font-size: 11.5px;" name="destelnuevo" />
															</td>
															<td>				
																<a href="#" onclick="javascript:agregartelefono('<%=datapersona%>')"><i class="demo-icon icon-floppy">&#xe809;</i><a>
															</td>
														</tr>
														</table>	
<%
desconectar
end if
%>