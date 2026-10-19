<%
' ==============================================================================
' PoliMusic - Arquitectura de Conexión y Helpers ADO
' Motor: Microsoft SQL Server (LocalDB)
' ==============================================================================

' Configuración de Codificación Global UTF-8
Session.CodePage = 65001
Response.CodePage = 65001
Response.CharSet = "UTF-8"

' ------------------------------------------------------------------------------
' Constantes ADO (ActiveX Data Objects) para VBScript
' ------------------------------------------------------------------------------
Const adInteger         = 3
Const adVarChar         = 200
Const adLongVarChar     = 201
Const adVarWChar        = 202
Const adDBTimeStamp     = 135
Const adDate            = 7
Const adBoolean         = 11

Const adParamInput      = 1
Const adParamOutput     = 2
Const adParamReturnValue = 4

Const adCmdText         = 1
Const adCmdStoredProc   = 4

Const adOpenForwardOnly = 0
Const adOpenKeyset      = 1
Const adOpenDynamic     = 2
Const adOpenStatic      = 3

Const adLockReadOnly    = 1
Const adLockPessimistic = 2
Const adLockOptimistic  = 3

' ------------------------------------------------------------------------------
' Cadena de Conexión Estricta requerida por la especificación del proyecto
' ------------------------------------------------------------------------------
Const DB_CONNECTION_STRING = "Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=BDD_PoliMusic;Integrated Security=True;TrustServerCertificate=True;"

' ------------------------------------------------------------------------------
' Función para obtener una conexión abierta a la Base de Datos
' ------------------------------------------------------------------------------
Function GetConnection()
    Dim conn, activeCs
    Set conn = Server.CreateObject("ADODB.Connection")
    
    On Error Resume Next
    ' Intento 1: Apertura directa con la cadena especificada
    conn.Open DB_CONNECTION_STRING

    ' Si el proveedor OLE DB subyacente requiere Provider o SSPI (estándar Windows para LocalDB)
    If Err.Number <> 0 Then
        Err.Clear
        activeCs = DB_CONNECTION_STRING
        ' En OLE DB, autenticación Windows es SSPI
        activeCs = Replace(activeCs, "Integrated Security=True", "Integrated Security=SSPI", 1, -1, 1)
        If InStr(1, activeCs, "Provider=", 1) = 0 Then
            activeCs = "Provider=MSOLEDBSQL;" & activeCs
        End If
        conn.Open activeCs

        If Err.Number <> 0 Then
            Err.Clear
            ' Fallback a SQLNCLI11 o SQLOLEDB con Integrated Security
            activeCs = Replace(activeCs, "Provider=MSOLEDBSQL;", "Provider=SQLNCLI11;", 1, -1, 1)
            conn.Open activeCs

            If Err.Number <> 0 Then
                Err.Clear
                ' Fallback con usuario SQL creado en BDD_PoliMusic.sql (usr_polimusic / Politecnica1)
                activeCs = "Provider=MSOLEDBSQL;Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=BDD_PoliMusic;User ID=usr_polimusic;Password=Politecnica1;TrustServerCertificate=True;"
                conn.Open activeCs

                If Err.Number <> 0 Then
                    Err.Clear
                    activeCs = "Provider=SQLOLEDB;Data Source=(localdb)\MSSQLLocalDB;Initial Catalog=BDD_PoliMusic;User ID=usr_polimusic;Password=Politecnica1;"
                    conn.Open activeCs
                End If
            End If
        End If
    End If
    On Error GoTo 0

    Set GetConnection = conn
End Function

' ------------------------------------------------------------------------------
' Función para cerrar y liberar objetos de forma segura
' ------------------------------------------------------------------------------
Sub CloseConnection(ByRef conn)
    On Error Resume Next
    If IsObject(conn) Then
        If conn.State = 1 Then conn.Close
        Set conn = Nothing
    End If
    On Error GoTo 0
End Sub

Sub CloseRecordset(ByRef rs)
    On Error Resume Next
    If IsObject(rs) Then
        If rs.State = 1 Then rs.Close
        Set rs = Nothing
    End If
    On Error GoTo 0
End Sub

' ------------------------------------------------------------------------------
' Helpers de Interfaz y Utilidades
' ------------------------------------------------------------------------------

' Resuelve rutas tipo ASP.NET "~/SongFiles/cancion.mp3" a rutas web relativas "SongFiles/cancion.mp3"
Function ResolveUrl(ByVal rawPath)
    If IsNull(rawPath) Or Trim(rawPath) = "" Then
        ResolveUrl = "img/default-cover.svg"
        Exit Function
    End If
    
    Dim p
    p = Trim(rawPath)
    If Left(p, 2) = "~/" Or Left(p, 2) = "~\ " Then
        p = Mid(p, 3)
    ElseIf Left(p, 1) = "/" Or Left(p, 1) = "~" Then
        p = Mid(p, 2)
    End If
    
    p = Replace(p, "\", "/")
    ResolveUrl = p
End Function

' Escapa caracteres HTML de manera segura evitando errores de valor nulo
Function SafeHtml(ByVal val)
    If IsNull(val) Or IsEmpty(val) Then
        SafeHtml = ""
    Else
        SafeHtml = Server.HTMLEncode(CStr(val))
    End If
End Function

' Formatea fechas para presentación estándar
Function FormatDateStr(ByVal val)
    If IsNull(val) Or IsEmpty(val) Or Not IsDate(val) Then
        FormatDateStr = "N/D"
    Else
        FormatDateStr = Day(val) & "/" & Month(val) & "/" & Year(val)
    End If
End Function

' Retorna el nombre legible del rol
Function FormatRoleName(ByVal userType)
    If IsNull(userType) Then
        FormatRoleName = "Desconocido"
    ElseIf CInt(userType) = 1 Then
        FormatRoleName = "Administrador"
    Else
        FormatRoleName = "Usuario EPN"
    End If
End Function

' Formatea el tipo de playlist
Function FormatPlaylistType(ByVal pType)
    If IsNull(pType) Then
        FormatPlaylistType = "Pública"
    ElseIf CInt(pType) = 1 Then
        FormatPlaylistType = "Privada"
    Else
        FormatPlaylistType = "Pública"
    End If
End Function
%>
