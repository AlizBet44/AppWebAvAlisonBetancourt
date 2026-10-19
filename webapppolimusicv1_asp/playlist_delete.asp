<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim playlistId
playlistId = Request.QueryString("id")

If Not IsNumeric(playlistId) Then
    Response.Redirect "playlists.asp"
    Response.End
End If

Dim conn, cmdCheck, rsCheck
Set conn = GetConnection()

' 1. Validar existencia y propiedad
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

' RBAC: Solo el dueño de la playlist o un userAdmin pueden eliminarla
If CLng(ownerId) <> CurrentUserId() And Not IsAdmin() Then
    CloseConnection conn
    Response.Redirect "playlists.asp?msg=unauthorized_delete"
    Response.End
End If

' 2. Eliminar canciones vinculadas en TBL_PLAYLIST_SONG
Dim cmdDelSongs
Set cmdDelSongs = Server.CreateObject("ADODB.Command")
cmdDelSongs.ActiveConnection = conn
cmdDelSongs.CommandType = adCmdText
cmdDelSongs.CommandText = "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?"
cmdDelSongs.Parameters.Append cmdDelSongs.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
cmdDelSongs.Execute()

' 3. Eliminar la playlist en TBL_PLAYLIST
Dim cmdDelPl
Set cmdDelPl = Server.CreateObject("ADODB.Command")
cmdDelPl.ActiveConnection = conn
cmdDelPl.CommandType = adCmdText
cmdDelPl.CommandText = "DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?"
cmdDelPl.Parameters.Append cmdDelPl.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
cmdDelPl.Execute()

CloseConnection conn
Response.Redirect "playlists.asp?msg=deleted"
%>
