<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim playlistId, songId, returnUrl
playlistId = Request("playlist_id")
songId     = Request("song_id")
returnUrl  = Request("return_url")

If returnUrl = "" Then returnUrl = "playlist_detail.asp?id=" & playlistId

If Not IsNumeric(playlistId) Or Not IsNumeric(songId) Then
    Response.Redirect returnUrl
    Response.End
End If

Dim conn, cmdCheck, rsCheck
Set conn = GetConnection()

' 1. Validar propiedad de la playlist o rol administrador
Set cmdCheck = Server.CreateObject("ADODB.Command")
cmdCheck.ActiveConnection = conn
cmdCheck.CommandType = adCmdText
cmdCheck.CommandText = "SELECT ID_USER FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?"
cmdCheck.Parameters.Append cmdCheck.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
Set rsCheck = cmdCheck.Execute()

If rsCheck.EOF Then
    CloseRecordset rsCheck
    CloseConnection conn
    Response.Redirect "playlists.asp"
    Response.End
End If

Dim ownerId
ownerId = rsCheck("ID_USER")
CloseRecordset rsCheck

If CLng(ownerId) <> CurrentUserId() And Not IsAdmin() Then
    CloseConnection conn
    Response.Redirect "playlists.asp?msg=unauthorized"
    Response.End
End If

' 2. Comprobar si la canción ya está en la playlist para evitar error de clave duplicada
Dim cmdExist, rsExist
Set cmdExist = Server.CreateObject("ADODB.Command")
cmdExist.ActiveConnection = conn
cmdExist.CommandType = adCmdText
cmdExist.CommandText = "SELECT COUNT(*) AS Total FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?"
cmdExist.Parameters.Append cmdExist.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
cmdExist.Parameters.Append cmdExist.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(songId))
Set rsExist = cmdExist.Execute()

If Not rsExist.EOF And CLng(rsExist("Total")) = 0 Then
    CloseRecordset rsExist

    ' 3. Insertar relación parametrizada
    Dim cmdInsert
    Set cmdInsert = Server.CreateObject("ADODB.Command")
    cmdInsert.ActiveConnection = conn
    cmdInsert.CommandType = adCmdText
    cmdInsert.CommandText = "INSERT INTO TBL_PLAYLIST_SONG (ID_PLAYLIST, ID_SONG) VALUES (?, ?)"
    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(songId))
    cmdInsert.Execute()
Else
    CloseRecordset rsExist
End If

CloseConnection conn

If InStr(returnUrl, "?") > 0 Then
    Response.Redirect returnUrl & "&msg=song_added"
Else
    Response.Redirect returnUrl & "?msg=song_added"
End If
%>
