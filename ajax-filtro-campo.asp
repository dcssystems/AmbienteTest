<%@ LANGUAGE = VBScript.Encode %>
<!--#include file=capa1.asp-->
<!--#include file=capa2.asp-->  
<option value ="0" >Seleccione un filtro</option>
<%
if session("codusuario")<>"" then
	conectar
if permisofacultad("dcs_admfacultad.asp") then
	id = obtener("elegido")

sql = "SELECT idfiltro, descripcion FROM Filtro WHERE TipoCampo = ( Select TipoCampo from Campaña_campo where IDCampañaCampo = " & id & " )"
consultar sql,RS

Do While Not  RS.EOF
%>
	<option value="<%=RS.Fields("idfiltro")%>" <% if idfiltro<>"" then%><% if RS.fields("idfiltro")=int(idfiltro) then%> selected<%end if%><%end if%>><%=RS.Fields("descripcion")%></option>
<%
RS.MoveNext
loop
RS.Close

end if
end if


%>