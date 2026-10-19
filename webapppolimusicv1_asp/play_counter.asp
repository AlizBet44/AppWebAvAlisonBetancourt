<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<%
' ==============================================================================
' PoliMusic - Endpoint AJAX para Incrementar Contador de Reproducciones
' ==============================================================================
Response.ContentType = "application/json"

Dim songId
songId = Request("id")

If IsNumeric(songId) Then
    On Error Resume Next
    Dim conn, cmd
    Set conn = GetConnection()
    Set cmd = Server.CreateObject("ADODB.Command")
    cmd.ActiveConnection = conn
    cmd.CommandType = adCmdText
    cmd.CommandText = "UPDATE TBL_SONG SET PLAYS = ISNULL(PLAYS, 0) + 1 WHERE ID_SONG = ?"
    cmd.Parameters.Append cmd.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(songId))
    cmd.Execute()
    
    CloseConnection conn
    On Error GoTo 0
    
    Response.Write "{""success"": true, ""id"":" & CLng(songId) & "}"
Else
    Response.Write "{""success"": false, ""error"": ""ID no válido""}"
End If
%>
