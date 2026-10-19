<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim playlistId, songId
playlistId = Request("playlist_id")
songId     = Request("song_id")

If Not IsNumeric(playlistId) Or Not IsNumeric(songId) Then
    Response.Redirect "playlists.asp"
    Response.End
End If

Dim conn, cmdCheck, rsCheck
Set conn = GetConnection()

' 1. Validar propiedad
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

' 2. Eliminar la canción de la lista
Dim cmdDelete
Set cmdDelete = Server.CreateObject("ADODB.Command")
cmdDelete.ActiveConnection = conn
cmdDelete.CommandType = adCmdText
cmdDelete.CommandText = "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ? AND ID_SONG = ?"
cmdDelete.Parameters.Append cmdDelete.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
cmdDelete.Parameters.Append cmdDelete.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(songId))
cmdDelete.Execute()

CloseConnection conn
Response.Redirect "playlist_detail.asp?id=" & playlistId & "&msg=song_removed"
%>
