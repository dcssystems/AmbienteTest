<% 
Response.addHeader "pragma", "no-cache"
Response.CacheControl = "Private"
' Selecciona una de las tres opciones siguientes
Response.Expires = -1441
Response.Expires = 0
Response.ExpiresAbsolute = #1/5/2000 12:12:12#
session.Timeout=600
Server.ScriptTimeout=999999999

'Se agrego estas variables para lograr hacer el registro de sesiones
Dim IP,Host,User

IP   = request.ServerVariables("REMOTE_ADDR")
Host = request.ServerVariables("REMOTE_HOST")
User = request.ServerVariables("REMOTE_USER")


%>
<script language="javascript">
	
	function global_popup_IWTSystem(ventana,ruta,nombre,propiedades)
	{
		if(ventana==null)
		{
		ventana=window.open(ruta,nombre,propiedades);	
		}
		else
		{
			if(ventana.closed)			
			{
			ventana=window.open(ruta,nombre,propiedades);	
			}
		}
		ventana.focus();	
		return ventana;			
	}
	function global_formatnumber_IWTSystem(Numero, Decimales) {

	pot = Math.pow(10,Decimales);
	num = parseInt(Numero * pot) / pot;
	nume = num.toString().split('.');

	entero = nume[0];
	decima = nume[1];

	if (decima != undefined) {
		fin = Decimales-decima.length; }
	else {
		decima = '';
		fin = Decimales; }

	for(i=0;i<fin;i++)
	  decima+=String.fromCharCode(48); 

    buffer="";
	marca=entero.length-1;
	chars=1;
	while(marca>=0){
	   if((chars%4)==0){
		  buffer=","+buffer;
	   }
	   buffer=entero.charAt(marca)+buffer;
	   marca--;
	   chars++;
	}
	num=buffer+'.'+decima;
	return num;
}
</script>
<%

Function obtener(objeto)
''Para obtener el dato de un formulario
	obtener=replace(request(objeto),"'","")
End Function


Function Dec(isHex) 
'Funcion que convierte un numero HexaDecimal a Decimal
	isHex = UCase(isHex) 
	If Len(isHex) <> 2 Then 
	Dec = -1 
	Exit Function 
	End If 
	Dim lsCaracter1,lsCaracter2,lnCaracter1,lnCaracter2 
	lsCaracter1 = Mid(isHex, 2, 1) 
	lsCaracter2 = Mid(isHex, 1, 1) 
	If (lsCaracter1 < "0" Or lsCaracter1 > "9") And (lsCaracter1 < "A" Or lsCaracter1 > "F") Then 
	Dec = -1 
	Exit Function 
	End If 
	If (lsCaracter2 < "0" Or lsCaracter2 > "9") And (lsCaracter2 < "A" Or lsCaracter2 > "F") Then 
	Dec = -1 
	Exit Function 
	End If 
	If (lsCaracter1 >= "0" And lsCaracter1 <= "9") Then 
	lnCaracter1 = Asc(lsCaracter1) - Asc("0") 
	ElseIf (lsCaracter1 >= "A" And lsCaracter1 <= "F") Then 
	lnCaracter1 = 10 + Asc(lsCaracter1) - Asc("A") 
	End If 
	If (lsCaracter2 >= "0" And lsCaracter2 <= "9") Then 
	lnCaracter2 = Asc(lsCaracter2) - Asc("0") 
	ElseIf (lsCaracter2 >= "A" And lsCaracter2 <= "F") Then 
	lnCaracter2 = 10 + Asc(lsCaracter2) - Asc("A") 
	End If 
	Dec = 16 * lnCaracter2 + lnCaracter1 
End Function 


Public Function Invertir(isCadena) 
	'funcion que Invierte la Cadena
	Dim lnLongitud 
	lnLongitud = Len(isCadena) 
	If lnLongitud = 0 Then 
	Invertir = "" 
	Exit Function 
	End If 
	If lnLongitud = 1 Then 
	Invertir = isCadena 
	Exit Function 
	End If 
	Invertir = Mid(isCadena, lnLongitud, 1) & Invertir(Mid(isCadena, 1, lnLongitud - 1)) 
End Function 

Function Desencriptar (iscadena) 
'Funcion que Desencripta
	if IsNull(iscadena) then
	iscadena=""
	end if
	Dim lnContador, lnLongitud, lnAscii,lsCaracterRnd, lsCaracterDes, lsCaracter, lsCadenaEnc, lsCadenaDes 
	lsCadenaEnc = Invertir(iscadena) 
	lnLongitud = Len(lsCadenaEnc) / 2 
	If lnLongitud Mod 2 <> 0 Then 
	Desencriptar = "" 
	Exit Function 
	End If 
	For lnContador = 1 To lnLongitud Step 2 
	lsCaracterRnd = Chr(dec(Mid(lsCadenaEnc, 2 * lnContador - 1, 2))) 
	lsCaracter = Chr(dec(Mid(lsCadenaEnc, 2 * lnContador + 1, 2))) 
	lnAscii = (Asc(lsCaracter) - Asc(lsCaracterRnd)) Mod 256 
	If lnAscii < 0 Then 
	lsCaracterDes = Chr(256 + lnAscii) 
	Else 
	lsCaracterDes = Chr(lnAscii) 
	End If 
	lsCadenaDes = lsCadenaDes + lsCaracterDes 
	Next 
	Desencriptar = lsCadenaDes 
End function 


Public Function Encriptar(isCadena) 
'Funcion que Encripta la Cadena
	if IsNull(isCadena) then
	isCadena=""
	end if
	Dim lnContador,lnLongitud,lnCaracterRnd,lnCaracterEnc,lnCaracter,lsCadenaEnc 
	Randomize 
	lnLongitud = Len(iscadena) 
	For lnContador = 1 To lnLongitud 
	lnCaracter = Asc(Mid(iscadena, lnContador, 1)) 
	lnCaracterRnd = Int(256 * Rnd()) 
	Do Until lnCaracterRnd <> 0 And _ 
	lnCaracterRnd <> 8 And _ 
	lnCaracterRnd <> 9 And _ 
	lnCaracterRnd <> 10 And _ 
	lnCaracterRnd <> 13 And _ 
	(lnCaracter + lnCaracterRnd) Mod 256 <> 0 And _ 
	(lnCaracter + lnCaracterRnd) Mod 256 <> 8 And _ 
	(lnCaracter + lnCaracterRnd) Mod 256 <> 9 And _ 
	(lnCaracter + lnCaracterRnd) Mod 256 <> 10 And _ 
	(lnCaracter + lnCaracterRnd) Mod 256 <> 13 
	lnCaracterRnd = Int(256 * Rnd()) 
	Loop 
	lnCaracterEnc = (lnCaracter + lnCaracterRnd) Mod 256 
	lsCadenaEnc = lsCadenaEnc + Right("00" + Hex(lnCaracterRnd), 2) + Right(">" + Hex(lnCaracterEnc), 2) 
	Next 
	if instr(Invertir(lsCadenaEnc),"<")>0 or instr(Invertir(lsCadenaEnc),">")>0 then
		Encriptar = Encriptar(isCadena)
	else
	    Encriptar = Invertir(lsCadenaEnc) 
	end if
End Function 
'''la forma de llamarlas : 
'''realizar tu select q traiga el user and pwd y preguntas 
'''If Not Rs.EOF Then 
'''Password=Desencriptar(rs("pwd")) 
'''end if 
'''******************** 
'''Para encriptar 
'''password = Encriptar(Pwd) 
'''Pwd es un campo q se recupera con request("Pwd") 
'''y luegos realizar tu insert con el password Encriptado. 


Function validarusuario(usr,pwd,mensajevalida)
''Para Validar un usuario
''3er Intento Bloqueo
	sql="select codusuario,clave,flagbloqueo,nombres,apepaterno,apematerno,Anexo, ClaveAnexo, GoUsuario, GoClave from Usuario where usuario='" & usr &"' and activo=1"
	consultar sql,RS
	if not RS.EOF then
		if pwd=Desencriptar(RS.Fields("clave")) and RS.Fields("flagbloqueo")<3 then
			session("codusuario")=RS.Fields("codusuario")
			session("codusuariodif")=RS.Fields("codusuario") & replace(Date(),"/","") & "h" & Hour(Now()) & "m" & Minute(Now()) & "s" & Second(Now())
			session("nombreusuario")=iif(IsNull(RS.Fields("nombres")),"",RS.Fields("nombres")) & ", " & iif(IsNull(RS.Fields("apepaterno")),"",RS.Fields("apepaterno")) & " " & iif(IsNull(RS.Fields("apematerno")),"",RS.Fields("apematerno"))
			session("codigoPais") 	  = "51"
			session("anexo")      	  = RS.Fields("Anexo")
			session("claveAnexo") 	  = RS.Fields("ClaveAnexo")
			session("goUsuario")  	  = RS.Fields("GoUsuario")
			session("goClaveUsuario") = RS.Fields("GoClave")
			session("ipsession")	  = IP			
			'session("telefono")
			'session("telefono")			
			sql="Update Usuario set flagbloqueo=0 from Usuario where usuario='" & usr &"' and activo=1"
			conn.execute sql				
			validarusuario=true
			sqlSession = "INSERT INTO Session_activa VALUES(" & session("codusuario") & ",getdate(),'"& IP & "', '" & session("anexo") & "', NULL, NULL)"
			conn.execute sqlSession
		else
			if RS.Fields("flagbloqueo")<3 then
				sql="Update Usuario set flagbloqueo=flagbloqueo + 1 from Usuario where usuario='" & usr &"' and activo=1"
				conn.execute sql				
				select case (2 - RS.Fields("flagbloqueo")) 
					case 0: mensajevalida="Cuenta bloqueada por reiterados intentos de ingreso fallido.\nComunicarse con el administrador del sistema."
					case 1:	mensajevalida="La clave ingresada no es válida.\nQueda " & (2 - RS.Fields("flagbloqueo")) & " intento antes de bloquear tu cuenta."
					case else mensajevalida="La clave ingresada no es válida.\nQuedan " & (2 - RS.Fields("flagbloqueo")) & " intentos antes de bloquear tu cuenta."
				end select
				expirarusuario
				validarusuario=false
			else	
				mensajevalida="Cuenta bloqueada por reiterados intentos de ingreso fallido.\nComunicarse con el administrador del sistema."
				expirarusuario
				validarusuario=false
			end if
		end if
	else
		mensajevalida="Ud. No tiene autorización para ingresar al Sistema"
		expirarusuario
		validarusuario=false
	end if
	RS.Close
End Function

Function expirarusuario()
conectar
sql="DELETE FROM Session_activa WHERE idSession=" & session("codusuario")
conn.execute sql

session("codusuario")     =""
session("codusuariodif")  =""
session("nombreusuario")  =""
session("codigoPais") 	  =""
session("anexo")      	  =""
session("claveAnexo") 	  =""
session("goUsuario")  	  =""
session("goClaveUsuario") =""
session("ipsession")      =""

desconectar
End Function

Function permisofacultad(pagina)
	sql="select count(*) + (select administrador from Usuario where codusuario=" & session("codusuario") & ") from Facultad A inner join PerfilFacultad B on A.codfacultad=B.codfacultad inner join UsuarioPerfil C on B.codperfil=C.codperfil where C.codusuario=" & session("codusuario") & " and C.activo=1 and pagina='" & pagina & "'"
	consultar sql,RS
	if RS.Fields(0)>0 then
		permisofacultad=true
	else
		permisofacultad=false
	end if
	RS.Close
End Function

Function passwordactual()
	sql="select clave from Usuario where codusuario=" & session("codusuario")
	consultar sql,RS
	if not RS.EOF then
		passwordactual=Desencriptar(RS.Fields("clave")) 
	end if
	RS.Close
End Function

Function actualizapassword(pwd)
sql="update Usuario set clave='" & Encriptar(pwdnew) & "' where codusuario=" & session("codusuario")
conn.execute sql
End Function

Function ExportarExcel(excel)
	nombrearch="export" & session("idusuario")
	Act_CFile = Server.MapPath ("exportados/" & nombrearch & ".xls")
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	if fso.FileExists(Act_CFile) then fso.DeleteFile(Act_CFile)
	Set fso=nothing
	Set Fcount = CreateObject("Scripting.FileSystemObject")
	Set CounterFile = Fcount.CreateTextFile(Act_CFile, True)
	CounterFile.WriteLine (excel)
	CounterFile.Close 
	Set Fcount = nothing
	Set CounterFile = nothing
	ExportarExcel=nombrearch
End Function

Function GenerarXML(nombre,cadenaxml)
	nombrearch=nombre & session("idusuario")
	Act_CFile = Server.MapPath ("exportados/" & nombrearch & ".xml")
	Set fso = Server.CreateObject("Scripting.FileSystemObject") 
	if fso.FileExists(Act_CFile) then fso.DeleteFile(Act_CFile)
	Set fso=nothing
	Set Fcount = CreateObject("Scripting.FileSystemObject")
	Set CounterFile = Fcount.CreateTextFile(Act_CFile, True,True)
	CounterFile.WriteLine (cadenaxml)
	CounterFile.Close 
	Set Fcount = nothing
	Set CounterFile = nothing
End Function

Function iif(i,j,k)
  If i Then iif = j Else iif = k
End Function


Function ConvNumerosALetras(ConvNumber, Cadena_Adicional)

Dim ParteEntera, ParteDecimal
Dim Decimales, Enteros
Dim numdigts
      
   Enteros = Int(ConvNumber)
   Decimales = Round(ConvNumber - Enteros, 2) * 100
   numdigIts = Len(Trim(CStr(Enteros)))

   If Len(CStr(Enteros)) > 12 Then
      ''ParteEntera = Format(Enteros, "#,##0")
      ParteEntera = FormatNumber(Enteros, 0)
   Else
      For i = 1 To numdigIts
         VALDIGIT = Mid(CStr(Enteros), numdigIts - i + 1, 1)
         RET = ""
         Select Case i
            Case 1, 4, 7, 10
               Select Case VALDIGIT
                  Case 0
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           RET = "Diez "
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           RET = "Veinte "
                        Else
                           RET = ""
                        End If
                     Else
                        RET = ""
                     End If
                  Case 1
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           RET = "Once"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           RET = "Veintiuno "
                        Else
                           RET = "Un"
                        End If
                     Else
                        RET = "Un"
                     End If
                  Case 2
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           RET = "Doce"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           RET = "Veintidos"
                        Else
                           RET = "Dos"
                        End If
                     Else
                        RET = "Dos"
                     End If
                  Case 3
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           RET = "Trece"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           RET = "Veintitres"
                        Else
                           RET = "Tres"
                        End If
                     Else
                        RET = "Tres"
                     End If
                  Case 4
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           RET = "Catorce"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           RET = "Veinticuatro"
                        Else
                           RET = "Cuatro"
                        End If
                     Else
                        RET = "Cuatro"
                     End If
                  Case 5
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Quince"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Veiticinco"
                        Else
                           RET = "Cinco"
                        End If
                     Else
                        RET = "Cinco"
                     End If
                  Case 6
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Dieciseis"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Veintiseis"
                        Else
                           RET = "Seis"
                        End If
                     Else
                        RET = "Seis"
                     End If
                  Case 7
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Diecisiete"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Veintisiete"
                        Else
                           RET = "Siete"
                        End If
                     Else
                        RET = "Siete"
                     End If
                  Case 8
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Dieciocho"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Veintiocho"
                        Else
                           RET = "Ocho"
                        End If
                     Else
                        RET = "Ocho"
                     End If
                  Case 9
                     If numdigIts > i Then
                        If Mid(CStr(Enteros), numdigIts - i, 1) = "1" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Diecinueve"
                        ElseIf Mid(CStr(Enteros), numdigIts - i, 1) = "2" Then
                           'RETVAL = Mid(RETVAL, 5)
                           RET = "Veintinueve"
                        Else
                           RET = "Nueve"
                        End If
                     Else
                        RET = "Nueve"
                     End If
               End Select
               If i = 4 Then
                  RET = RET & " Mil "
               ElseIf i = 7 Then
                  RET = RET & " Millones "
               ElseIf i = 10 Then
                  RET = RET & " Mil "
               End If
            Case 2, 5, 8, 11
               Select Case VALDIGIT
                  Case 0
                     RET = ""
                  Case 1
                     'RET = "DIEZ"
                  Case 2
                     'If Mid(CStr(Enteros), numdigIts - I + 2, 1) = "0" Then
                     '   RET = "Veinte"
                     'Else
                     '   RET = "Veinti"
                     'End If
                  Case 3
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                        RET = "Treinta"
                     Else
                        RET = "Treinta y "
                     End If
                  Case 4
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                       RET = "Cuarenta"
                     Else
                       RET = "Cuarenta y "
                     End If
                  Case 5
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                        RET = "Cincuenta "
                     Else
                        RET = "Cincuenta y "
                     End If
                  Case 6
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                        RET = "Sesenta "
                     Else
                        RET = "Sesenta y "
                     End If
                  Case 7
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                        RET = "Setenta "
                     Else
                        RET = "Setenta y "
                     End If
                  Case 8
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                        RET = "Ochenta "
                     Else
                        RET = "Ochenta y "
                     End If
                  Case 9
                     If Mid(CStr(Enteros), numdigIts - i + 2, 1) = "0" Then
                        RET = "Noventa "
                     Else
                        RET = "Noventa y "
                     End If
               End Select
            Case 3, 6, 9, 12
               Select Case VALDIGIT
                  Case 0
                     RET = ""
                  Case 1
                     RET = "Ciento "
                  Case 2
                     RET = "Doscientos "
                  Case 3
                     RET = "Trescientos "
                  Case 4
                     RET = "Cuatrocientos "
                  Case 5
                     RET = "Quinientos "
                  Case 6
                     RET = "Seiscientos "
                  Case 7
                     RET = "Setecientos "
                  Case 8
                     RET = "Ochocientos "
                  Case 9
                     RET = "Novecientos "
               End Select
         End Select
         RETVAL = RET & RETVAL
         'Debug.Print RETVAL
      Next
   End If
   
   
   If Decimales = 0 Then
	  ParteDecimal = " con 00/100 " & Cadena_Adicional
   Else
	  if Decimales < 10 Then	
		ParteDecimal = " con 0" & Decimales & "/100 " & Cadena_Adicional
      else
		ParteDecimal = " con " & Decimales & "/100 " & Cadena_Adicional
      end if      
   End If

   'If Right(RETVAL, 2) <> "Y " Then
   '   RETVAL = RETVAL & " Y "
   'End If
   RETVAL = "SON:  " & RETVAL & ParteDecimal
   ConvNumerosALetras = RETVAL
End Function						
%>



