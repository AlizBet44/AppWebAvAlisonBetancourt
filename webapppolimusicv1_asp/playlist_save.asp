<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

If Request.ServerVariables("REQUEST_METHOD") <> "POST" Then
    Response.Redirect "playlists.asp"
    Response.End
End If

Dim action, title, pType, photo, playlistId
action     = Request.Form("action")
title      = Trim(Request.Form("title"))
pType      = Request.Form("type")
photo      = Trim(Request.Form("photo"))
playlistId = Request.Form("id_playlist")

If title = "" Then
    Response.Redirect "playlists.asp?msg=error_empty_title"
    Response.End
End If

If Not IsNumeric(pType) Then pType = 0
If photo = "" Then photo = "~/playlistImages/default.jpg"

Dim conn
Set conn = GetConnection()

If action = "create" Then
    ' INSERT PARAMETRIZADO CON ADODB.Command
    Dim cmdInsert
    Set cmdInsert = Server.CreateObject("ADODB.Command")
    cmdInsert.ActiveConnection = conn
    cmdInsert.CommandType = adCmdText
    cmdInsert.CommandText = "INSERT INTO TBL_PLAYLIST (ID_USER, TITLE, CREATION_DATE, TYPE, PHOTO) VALUES (?, ?, GETDATE(), ?, ?)"

    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@ID_USER", adInteger, adParamInput, , CurrentUserId())
    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@TITLE", adVarChar, adParamInput, 50, title)
    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@TYPE", adInteger, adParamInput, , CInt(pType))
    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@PHOTO", adVarChar, adParamInput, 255, photo)

    cmdInsert.Execute()
    CloseConnection conn

    Response.Redirect "playlists.asp?msg=created"
    Response.End

ElseIf action = "edit" Then
    If Not IsNumeric(playlistId) Then
        CloseConnection conn
        Response.Redirect "playlists.asp"
        Response.End
    End If

    ' Validar Permiso (Solo dueño o Administrador)
    Dim cmdCheck, rsCheck
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

    ' UPDATE PARAMETRIZADO
    Dim cmdUpdate
    Set cmdUpdate = Server.CreateObject("ADODB.Command")
    cmdUpdate.ActiveConnection = conn
    cmdUpdate.CommandType = adCmdText
    cmdUpdate.CommandText = "UPDATE TBL_PLAYLIST SET TITLE = ?, TYPE = ?, PHOTO = ? WHERE ID_PLAYLIST = ?"

    cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@TITLE", adVarChar, adParamInput, 50, title)
    cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@TYPE", adInteger, adParamInput, , CInt(pType))
    cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@PHOTO", adVarChar, adParamInput, 255, photo)
    cmdUpdate.Parameters.Append cmdUpdate.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))

    cmdUpdate.Execute()
    CloseConnection conn

    Response.Redirect "playlist_detail.asp?id=" & playlistId & "&msg=updated"
    Response.End
End If

CloseConnection conn
Response.Redirect "playlists.asp"
%>
