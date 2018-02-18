<% 
Const adOpenStatic = 3
Const adLockReadOnly = 1
Const adCmdText = &H0001

'---- ConnectModeEnum Values ---- 
Const adModeUnknown = 0 
Const adModeRead = 1 
Const adModeWrite = 2 
Const adModeReadWrite = 3 
Const adModeShareDenyRead = 4 
Const adModeShareDenyWrite = 8 
Const adModeShareExclusive = &Hc 
Const adModeShareDenyNone = &H10 

'---- IsolationLevelEnum Values ---- 
Const adXactUnspecified = &Hffffffff 
Const adXactChaos = &H00000010 
Const adXactReadUncommitted = &H00000100 
Const adXactBrowse = &H00000100 
Const adXactCursorStability = &H00001000 
Const adXactReadCommitted = &H00001000 
Const adXactRepeatableRead = &H00010000 
Const adXactSerializable = &H00100000 
Const adXactIsolated = &H00100000 

Const adUseClient = 3

'---------------------Constantes para la Web----------------------------
Const TITLE = ".::DIRCON | Sistema web de gestión de cobranzas::."

'---------------------Fin Constantes para la Web------------------------



Dim enlace,conn,RS,RS1,RS2,RS3
Dim conn_server,conn_uid,conn_pwd,conn_database

set conn = server.CreateObject("ADODB.Connection")
set RS = server.CreateObject("ADODB.Recordset")
set RS1 = server.CreateObject("ADODB.Recordset")
set RS2 = server.CreateObject("ADODB.Recordset")
set RS3 = server.CreateObject("ADODB.Recordset")
set RS4 = server.CreateObject("ADODB.Recordset")
conn_server="192.168.1.7"
conn_uid="sa"
conn_pwd="d1rc0n$$"
conn_database="DataCRMDirconTest"
enlace="driver={SQL Server};server=" & conn_server & ";uid=" & conn_uid & ";pwd=" & conn_pwd & ";database=" & conn_database

Function conectar()
    conn.IsolationLevel = adXactReadUncommitted	
    conn.Mode = adModeReadWrite
    conn.CursorLocation = adUseClient
	conn.Open enlace
	conn.CommandTimeout=180
End Function

Function desconectar()
	conn.Close
End Function

Function consultar(consulta,registro)
registro.Open consulta,conn
End Function
%>



