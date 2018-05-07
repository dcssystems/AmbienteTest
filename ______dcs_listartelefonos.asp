<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<%
 
if session("codusuario")<>"" then
conectar
datapersona = request("datapersona")
idcampana = request("idcampana")

%>
<table class="tabinterna"  id="tabinterna_telf">
														<tr class="cabecera-orange">
															<td colspan="8">Tel&eacute;fonos
															</td>
														</tr>
														<tr class="fondo-red">
															<td class="text-withe">Tipo</td>
															<td class="text-withe">Prf</td>
															<td class="text-withe">N&uacute;mero</td>
															<td class="text-withe">Ext</td>
															<td class="text-withe">Descripci&oacute;n</td>
															<td class="text-withe"></td>
															<td class="text-withe"></td>
														</tr>
														<%
														if datapersona <> "" then

														varcolor = 0
															
														sql = "select a.IDCampañaPersonaTelefono,(select descripcion from TipoTelefono where IDTipoTelefono = a.IDTipoTelefono) as Tipo,a.Prefijo,a.Numero,a.Extension,a.Descripcion, a.Usuarioregistra, a.Enriquecido, a.IDTipoTelefono from Campaña_Persona_Telefono a where IDCampañaPersona =" & datapersona

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
															<td style="background: #a42627; text-align: center;"><a href="#" onclick="javascript:creargestion('<%=datapersona%>','<%=idcampana%>','<%=RS4.Fields("Numero")%>','Llamando...');"><div><i id="telef<%=RS4.Fields("Numero")%>" class="demo-icon6 icon-phone-circled telefono-inactivo">&#xe822;</i></div></a></td>
															<td  style="background: #a42627; text-align: center;"><a href="#" onclick="javascript:creargestion('<%=datapersona%>','<%=idcampana%>','<%=RS4.Fields("Numero")%>','Ingresando datos...');"><i id="addges<%=RS4.Fields("Numero")%>" class="demo-icon6 icon-plus-squared telefono-inactivo">&#xf0fe;</i></a></td>
															<td  style="background: #a42627; text-align: center;"><% if RS4.Fields("Usuarioregistra") = session("codUsuario") and  RS4.Fields("Enriquecido") = "1" then%><a href="#" onclick="javascript:editartelefono('<%=RS4.Fields("IDTipoTelefono")%>','<%=RS4.Fields("Numero")%>','<%=RS4.Fields("Extension")%>','<%=RS4.Fields("Descripcion")%>','<%=RS4.Fields("IDCampañaPersonaTelefono")%>')"><i id="Edtelf<%=RS4.Fields("Numero")%>" class="demo-icon6 icon-pencil-squared telefono-inactivo">&#xf14b;</i></a><%end if%></td>
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
																<a href="#" onclick="javascript:agregartelefono('<%=datapersona%>','<%=idcampana%>')"><i class="demo-icon icon-floppy">&#xe809;</i><a>
															</td>
															<td colspan="2">
																<input type="hidden" name="idpertel" id="idpertel" value="" />
																<p id="text-nuevo-edit" style="color:#E9F7F7;"></p>
															</td>
														</tr>
														</table>


<%
desconectar
end if
%>