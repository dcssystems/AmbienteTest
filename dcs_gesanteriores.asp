<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<%
if session("codusuario")<>"" then
conectar
datapersona = Request("datapersona")
%>
									<table class="tabinterna" name="gesanteriores" id="gesanteriores">
									<tr class="cabecera-orange">
										<td class="text-withe" colspan="6" style="text-align: center; font-weight: bold;">
											Gestiones Anteriores
										</td>
									</tr>
									<tr class="fondo-red">
										<td class="text-withe">Gestor</td>
										<td class="text-withe">Acción</td>
										<td class="text-withe">Fecha</td>
										<td class="text-withe">Telefono</td>
										<td class="text-withe">Gestion</td>
										<td class="text-withe">Comentario</td>
									</tr>
									<%
									if datapersona <> "" then

										sql = "select (Select Usuario from Usuario where codUsuario = UsuarioEjecutor) as Gestor,(select Descripcion from TipoAccion where IDTipoAccion = a.IDTipoAccion) as Accion,FechaHoraFinGestion as fechagestion,(Select Numero from Campaña_persona_telefono where IDCampañaPersonaTelefono = a.IDCampañaPersonaTelefono )  as Telefono,(Select Descripcion from Gestion where IDGestion = a.IDGestion) as Respuesta,Comentario from Campaña_Persona_Accion a where a.IDCampañaPersona = " & datapersona & "order by FechaHoraFinGestion desc"

										consultar sql,RS4
										varcolor = 0
										DO While not RS4.EOF

									%>
									<tr class="fondo-red <% IF varcolor	 = 0 Then %> fondo-blanco <% Else %> fondo-rojo <% End IF %>" >
										<td><%=RS4.fields("Gestor")%></td>
										<td><%=RS4.fields("Accion")%></td>
										<td><%=RS4.fields("fechagestion")%></td>
										<td><%=RS4.fields("Telefono")%></td>
										<td><%=RS4.fields("Respuesta")%></td>
										<td><%=RS4.fields("Comentario")%></td>
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

								</table>

									<%					
							
					desconectar
					

end if
%>