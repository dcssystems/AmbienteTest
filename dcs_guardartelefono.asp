<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<% 
if session("codusuario")<>"" then
	conectar

				idpersona = request("idpersona")
				tipotelf = request("tipotelf")
				numero = request("numero")
				extension = request("extension")
				descripcion = request("descripcion")
				res =""
				res2 =""
				

				sql = "Select count(IDCampañaPersonaTelefono) as 'Num' from Campaña_Persona_Telefono where Numero ='" & numero & "' and IDCampañaPersona =" & idpersona

				consultar sql,RS6

				res = RS6.fields("Num")

				RS6.Close

				if res > 0 then

				response.write "2" ' ya existe este telefono en la base de datos para esta persona

				else


				sql = "INSERT INTO [dbo].[Campaña_Persona_Telefono] ([IDTipoTelefono] ,[Prefijo] ,[Numero] ,[Extension] ,[Descripcion] ,[UsuarioRegistra] " & chr(10) & _
				 		",[FechaRegistra] ,[UsuarioModifica] ,[FechaModifica] ,[Enriquecido] ,[IDCampañaPersona] ,[DatoReferido]) " & chr(10) & _
				     "VALUES " & chr(10) & _
					   "( '" & tipotelf & "','51' ,'" & numero & "' ,'" & extension & "', '" & descripcion & "' , '" & session("codusuario") & "' " & chr(10) & _
				           ",getdate() ,NULL  ,NULL   ,1     ,'" & idpersona & "' ,0)  "
				         


				  conn.Execute sql

				  sql=   "Select Count(SCOPE_IDENTITY()) as Id"
				  consultar sql,RS6

				  res2 = RS6.fields("Id")
				
				 RS6.Close

											if res2 > 0 then

											  response.write "1" 'se agrego telefono correctamente.
											else

											  response.write "0" 'ocurrio un error.
											 
											end if
											' response.write "se agrego telefono correctamente"


				desconectar
			end if
else
response.write "se cerro la session"
			
end if
%>