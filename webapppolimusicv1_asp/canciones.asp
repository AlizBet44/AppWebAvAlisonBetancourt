<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
' ==============================================================================
' PoliMusic - Módulo de Canciones (canciones.asp)
' Evidencia Académica y Marca Personal: alison.betancourt
' Control RBAC: Administrador (CRUD Total) vs Usuario Normal (Solo Lectura y Play)
' ==============================================================================

' 1. Verificación de Autenticación Obligatoria
RequireLogin()

' 2. Lectura Obligatoria del Nombre del Usuario desde la Variable de Sesión (login.asp)
Dim sessionUsername, sessionUserType, sessionUserPhoto, isAdminUser
sessionUsername = Session("Username")
If IsNull(sessionUsername) Or Trim(sessionUsername & "") = "" Then
    sessionUsername = Session("usuario")
End If
If IsNull(sessionUsername) Or Trim(sessionUsername & "") = "" Then
    sessionUsername = CurrentUsername()
End If

sessionUserType  = Session("UserType")
sessionUserPhoto = CurrentUserPhoto()
isAdminUser      = IsAdmin()

Dim conn
Set conn = GetConnection()

Dim infoMsg, errorMsg
infoMsg  = ""
errorMsg = ""

' ------------------------------------------------------------------------------
' 3. Procesamiento de Operaciones CRUD (Estrictamente protegido para Administrador)
' ------------------------------------------------------------------------------
Dim reqMethod, action
reqMethod = Request.ServerVariables("REQUEST_METHOD")
action    = Request("action")

If reqMethod = "POST" Or action = "delete" Then
    If Not isAdminUser Then
        errorMsg = "Acceso denegado: No tienes permisos para realizar operaciones CRUD sobre el catálogo."
    Else
        ' OPERACIÓN CRUD: AGREGAR NUEVA CANCIÓN
        If reqMethod = "POST" And action = "create" Then
            Dim newName, newPath, newArtistId, newAlbumId
            newName     = Trim(Request.Form("song_name"))
            newPath     = Trim(Request.Form("song_path"))
            newArtistId = Request.Form("artist_id")
            newAlbumId  = Request.Form("album_id")

            If newName = "" Or newPath = "" Then
                errorMsg = "El nombre y la ruta de la canción son obligatorios."
            Else
                On Error Resume Next
                Dim cmdInsSong, rsNewSongId
                Set cmdInsSong = Server.CreateObject("ADODB.Command")
                cmdInsSong.ActiveConnection = conn
                cmdInsSong.CommandType = adCmdText
                cmdInsSong.CommandText = "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, 0); SELECT SCOPE_IDENTITY() AS NewId;"
                cmdInsSong.Parameters.Append cmdInsSong.CreateParameter("@SONG_NAME", adVarChar, adParamInput, 50, newName)
                cmdInsSong.Parameters.Append cmdInsSong.CreateParameter("@SONG_PATH", adVarChar, adParamInput, 255, newPath)
                Set rsNewSongId = cmdInsSong.Execute()

                If Err.Number <> 0 Then
                    errorMsg = "Error al registrar la canción: " & Err.Description
                    Err.Clear
                Else
                    Dim createdSongId
                    createdSongId = rsNewSongId("NewId")
                    CloseRecordset rsNewSongId

                    ' Asociar Artista
                    If IsNumeric(newArtistId) And CLng(newArtistId) > 0 Then
                        Dim cmdArt
                        Set cmdArt = Server.CreateObject("ADODB.Command")
                        cmdArt.ActiveConnection = conn
                        cmdArt.CommandType = adCmdText
                        cmdArt.CommandText = "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)"
                        cmdArt.Parameters.Append cmdArt.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(createdSongId))
                        cmdArt.Parameters.Append cmdArt.CreateParameter("@ID_ARTIST", adInteger, adParamInput, , CLng(newArtistId))
                        cmdArt.Execute()
                    End If

                    ' Asociar Álbum
                    If IsNumeric(newAlbumId) And CLng(newAlbumId) > 0 Then
                        Dim cmdAlb
                        Set cmdAlb = Server.CreateObject("ADODB.Command")
                        cmdAlb.ActiveConnection = conn
                        cmdAlb.CommandType = adCmdText
                        cmdAlb.CommandText = "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)"
                        cmdAlb.Parameters.Append cmdAlb.CreateParameter("@ID_ALBUM", adInteger, adParamInput, , CLng(newAlbumId))
                        cmdAlb.Parameters.Append cmdAlb.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(createdSongId))
                        cmdAlb.Execute()
                    End If

                    infoMsg = "¡Canción '" & SafeHtml(newName) & "' añadida exitosamente al catálogo!"
                End If
                On Error GoTo 0
            End If

        ' OPERACIÓN CRUD: EDITAR CANCIÓN
        ElseIf reqMethod = "POST" And action = "edit" Then
            Dim editId, editName, editPath, editArtistId, editAlbumId
            editId       = Request.Form("id_song")
            editName     = Trim(Request.Form("song_name"))
            editPath     = Trim(Request.Form("song_path"))
            editArtistId = Request.Form("artist_id")
            editAlbumId  = Request.Form("album_id")

            If Not IsNumeric(editId) Or editName = "" Or editPath = "" Then
                errorMsg = "Parámetros incompletos para actualizar la canción."
            Else
                On Error Resume Next
                Dim cmdUpdSong
                Set cmdUpdSong = Server.CreateObject("ADODB.Command")
                cmdUpdSong.ActiveConnection = conn
                cmdUpdSong.CommandType = adCmdText
                cmdUpdSong.CommandText = "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ? WHERE ID_SONG = ?"
                cmdUpdSong.Parameters.Append cmdUpdSong.CreateParameter("@SONG_NAME", adVarChar, adParamInput, 50, editName)
                cmdUpdSong.Parameters.Append cmdUpdSong.CreateParameter("@SONG_PATH", adVarChar, adParamInput, 255, editPath)
                cmdUpdSong.Parameters.Append cmdUpdSong.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(editId))
                cmdUpdSong.Execute()

                ' Actualizar relación con Artista
                conn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & CLng(editId)
                If IsNumeric(editArtistId) And CLng(editArtistId) > 0 Then
                    Dim cmdUpdArt
                    Set cmdUpdArt = Server.CreateObject("ADODB.Command")
                    cmdUpdArt.ActiveConnection = conn
                    cmdUpdArt.CommandType = adCmdText
                    cmdUpdArt.CommandText = "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)"
                    cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(editId))
                    cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@ID_ARTIST", adInteger, adParamInput, , CLng(editArtistId))
                    cmdUpdArt.Execute()
                End If

                ' Actualizar relación con Álbum
                conn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & CLng(editId)
                If IsNumeric(editAlbumId) And CLng(editAlbumId) > 0 Then
                    Dim cmdUpdAlb
                    Set cmdUpdAlb = Server.CreateObject("ADODB.Command")
                    cmdUpdAlb.ActiveConnection = conn
                    cmdUpdAlb.CommandType = adCmdText
                    cmdUpdAlb.CommandText = "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)"
                    cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@ID_ALBUM", adInteger, adParamInput, , CLng(editAlbumId))
                    cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(editId))
                    cmdUpdAlb.Execute()
                End If

                infoMsg = "Canción actualizada exitosamente."
                On Error GoTo 0
            End If

        ' OPERACIÓN CRUD: ELIMINAR CANCIÓN
        ElseIf action = "delete" Then
            Dim delId
            delId = Request("id")
            If IsNumeric(delId) Then
                On Error Resume Next
                Dim cmdDel
                Set cmdDel = Server.CreateObject("ADODB.Command")
                cmdDel.ActiveConnection = conn
                cmdDel.CommandType = adCmdText
                cmdDel.CommandText = "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?; " & _
                                     "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?; " & _
                                     "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?; " & _
                                     "DELETE FROM TBL_SONG WHERE ID_SONG = ?;"
                cmdDel.Parameters.Append cmdDel.CreateParameter("@P1", adInteger, adParamInput, , CLng(delId))
                cmdDel.Parameters.Append cmdDel.CreateParameter("@P2", adInteger, adParamInput, , CLng(delId))
                cmdDel.Parameters.Append cmdDel.CreateParameter("@P3", adInteger, adParamInput, , CLng(delId))
                cmdDel.Parameters.Append cmdDel.CreateParameter("@P4", adInteger, adParamInput, , CLng(delId))
                cmdDel.Execute()

                If Err.Number = 0 Then
                    infoMsg = "La canción fue eliminada del catálogo satisfactoriamente."
                Else
                    errorMsg = "Error al eliminar: " & Err.Description
                    Err.Clear
                End If
                On Error GoTo 0
            End If
        End If
    End If
End If

' ------------------------------------------------------------------------------
' 4. Consultas de Datos: Canciones, Artistas y Álbumes
' ------------------------------------------------------------------------------
Dim rsSongs
Set rsSongs = conn.Execute("SELECT " & _
    "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 p.ID_ARTIST FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG) AS ID_ARTIST, " & _
    "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
    " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME, " & _
    "(SELECT TOP 1 tl.ID_ALBUM FROM TBL_TRACKLIST tl WHERE tl.ID_SONG = s.ID_SONG) AS ID_ALBUM, " & _
    "(SELECT TOP 1 alb.TITLE_ALBUM " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_NAME, " & _
    "(SELECT TOP 1 alb.ALBUM_IMAGE_PATH " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE " & _
    "FROM TBL_SONG s ORDER BY s.ID_SONG DESC")

' Listas auxiliares para desplegables (si es Administrador)
Dim rsArtists, rsAlbums
If isAdminUser Then
    Set rsArtists = conn.Execute("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME ASC")
    Set rsAlbums  = conn.Execute("SELECT ID_ALBUM, TITLE_ALBUM FROM TBL_ALBUM ORDER BY TITLE_ALBUM ASC")
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Song List • Catálogo de Canciones • PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/canciones.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <!-- Notificaciones Dinámicas -->
        <% If infoMsg <> "" Then %>
          <div class="alert alert-success">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
            <span><%=SafeHtml(infoMsg)%></span>
          </div>
        <% End If %>

        <% If errorMsg <> "" Then %>
          <div class="alert alert-danger">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
            <span><%=SafeHtml(errorMsg)%></span>
          </div>
        <% End If %>

        <!-- ====================================================================
             REGLA: Saludo Dinámico + Evidencia y Marca Personal (alison.betancourt)
             ==================================================================== -->
        <section class="canciones-hero">
          <div class="hero-left">
            <img src="<%=sessionUserPhoto%>" alt="Avatar Usuario" class="user-avatar-large" onerror="this.src='UserImages/default.jpg'">
            <div class="hero-text-block">
              <div class="hero-welcome-badge">
                <span class="evidence-dot"></span>
                <span>Plataforma Multimedia PoliMusic</span>
              </div>
              <!-- Saludo dinámico obligatorio con Session("Username") -->
              <h1 class="hero-dynamic-greeting">
                ¡Bienvenida, <span class="highlight-user"><%=SafeHtml(sessionUsername)%></span>!
              </h1>
              <div class="hero-subinfo">
                <% If isAdminUser Then %>
                  <span class="role-pill role-pill-admin" title="Permisos Totales de Administración">
                    <svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.11v8.8z"/></svg>
                    Rol: Administrador (CRUD Habilitado)
                  </span>
                <% Else %>
                  <span class="role-pill role-pill-normal" title="Acceso de Usuario Politécnico">
                    <svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 3c1.66 0 3 1.34 3 3s-1.34 3-3 3-3-1.34-3-3 1.34-3 3-3zm0 14.2c-2.5 0-4.71-1.28-6-3.22.03-1.99 4-3.08 6-3.08 1.99 0 5.97 1.09 6 3.08-1.29 1.94-3.5 3.22-6 3.22z"/></svg>
                    Rol: Usuario Normal (Modo Escucha)
                  </span>
                <% End If %>
              </div>
            </div>
          </div>

          <div class="hero-actions">
            <!-- Botón de Cierre de Sesión claramente visible -->
            <a href="logout.asp" class="btn-logout-header" title="Cerrar la sesión actual">
              <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor">
                <path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4V5z"/>
              </svg>
              <span>Logout / Salir</span>
            </a>
          </div>
        </section>

        <!-- ====================================================================
             REGLA: Encabezado de Listado con el texto exacto "Song List"
             ==================================================================== -->
        <div class="catalog-section-head">
          <div class="catalog-title-group">
            <span class="catalog-tagline">
              <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/></svg>
              Catálogo Musical
            </span>
            <!-- Requisito Estricto: Texto "Song List" -->
            <h2 class="catalog-main-heading">
              Song List <span class="heading-accent">• Listado de Canciones</span>
            </h2>
            <p class="catalog-sub-heading">Explora y reproduce tus canciones en vivo con el reproductor HTML5 integrado</p>
          </div>

          <div class="catalog-controls">
            <!-- Buscador Dinámico en Vivo de la Tabla -->
            <div class="table-search-box">
              <svg viewBox="0 0 24 24"><path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>
              <input type="text" id="songFilterInput" class="table-search-input" placeholder="Filtrar por título, artista o álbum..." onkeyup="filterSongsTable()">
            </div>

            <!-- REGLA: Control de Roles en la Interfaz:
                 Solo si el usuario logueado es Administrador se renderizan opciones para CRUD -->
            <% If isAdminUser Then %>
              <button type="button" class="btn btn-primary" onclick="openAddSongModal()" id="btnOpenAddSong">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
                <span>Agregar Canción (Admin)</span>
              </button>
            <% End If %>
          </div>
        </div>

        <!-- ====================================================================
             REGLA: Tabla de Datos Estilizada + Reproducción Multimedia Integrada
             ==================================================================== -->
        <div class="canciones-table-card">
          <div class="canciones-table-container">
            <table class="custom-songs-table" id="tableSongsList">
              <thead>
                <tr>
                  <th class="col-center" style="width: 44px;">#</th>
                  <th>Título de la Canción</th>
                  <th>Artista</th>
                  <th>Álbum</th>
                  <th class="col-right" style="width: 100px;">Reproducciones</th>
                  <!-- REGLA: Columna de Reproducción Multimedia con <audio controls> -->
                  <th class="col-center" style="width: 270px;">Reproductor Multimedia</th>
                  <!-- REGLA: Columna CRUD Solo para Administrador -->
                  <% If isAdminUser Then %>
                    <th class="col-center" style="width: 150px;">Acciones (Admin)</th>
                  <% End If %>
                </tr>
              </thead>
              <tbody>
                <%
                Dim rowIdx, totalCount
                rowIdx = 1
                totalCount = 0

                If Not rsSongs.EOF Then
                    Do While Not rsSongs.EOF
                        totalCount = totalCount + 1
                        Dim songCover, songAudio, songArtist, songAlbum, songId, songName, songPath, songPlays
                        songId    = rsSongs("ID_SONG")
                        songName  = rsSongs("SONG_NAME")
                        songPath  = rsSongs("SONG_PATH")
                        songPlays = rsSongs("PLAYS")
                        
                        songCover  = ResolveUrl(rsSongs("ALBUM_IMAGE"))
                        songAudio  = ResolveUrl(songPath)
                        
                        songArtist = rsSongs("ARTIST_NAME")
                        If IsNull(songArtist) Or Trim(songArtist) = "" Then songArtist = "Artista PoliMusic"
                        
                        songAlbum  = rsSongs("ALBUM_NAME")
                        If IsNull(songAlbum) Or Trim(songAlbum) = "" Then songAlbum = "Sencillo"
                %>
                  <tr class="song-table-row" id="rowSong-<%=songId%>">
                    <!-- # Index -->
                    <td class="td-index"><%=rowIdx%></td>

                    <!-- Título y Portada -->
                    <td>
                      <div class="song-identity-cell">
                        <img src="<%=songCover%>" alt="Cover" class="song-cover-thumb" onerror="this.src='img/default-cover.svg'">
                        <div class="song-texts">
                          <span class="song-main-name"><%=SafeHtml(songName)%></span>
                          <span class="song-path-sub"><%=SafeHtml(songPath)%></span>
                        </div>
                      </div>
                    </td>

                    <!-- Artista -->
                    <td>
                      <span class="pill-artist">
                        <svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
                        <%=SafeHtml(songArtist)%>
                      </span>
                    </td>

                    <!-- Álbum -->
                    <td>
                      <span class="pill-album">
                        <svg viewBox="0 0 24 24" width="13" height="13" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 14.5c-2.49 0-4.5-2.01-4.5-4.5S9.51 7.5 12 7.5s4.5 2.01 4.5 4.5-2.01 4.5-4.5 4.5zm0-5.5c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>
                        <%=SafeHtml(songAlbum)%>
                      </span>
                    </td>

                    <!-- Reproducciones -->
                    <td class="td-plays-badge" id="playsSpan-<%=songId%>">
                      <%=songPlays%>
                    </td>

                    <!-- REGLA OBLIGATORIA: Reproductor Multimedia con <audio controls> HTML5 -->
                    <td class="td-player-cell">
                      <div class="inline-audio-container">
                        <audio controls preload="none" class="native-audio-player" data-song-id="<%=songId%>">
                          <source src="<%=songAudio%>" type="audio/mpeg">
                          Tu navegador no soporta el reproductor de audio HTML5.
                        </audio>
                      </div>
                    </td>

                    <!-- REGLA OBLIGATORIA: Control de Roles en la Interfaz (Admin CRUD) -->
                    <% If isAdminUser Then %>
                      <td class="td-actions-cell">
                        <div class="actions-btn-group">
                          <!-- Botón Editar -->
                          <button type="button" class="btn-crud btn-crud-edit" title="Editar Canción"
                                  data-id="<%=songId%>"
                                  data-name="<%=SafeHtml(songName)%>"
                                  data-path="<%=SafeHtml(songPath)%>"
                                  data-artist="<%=rsSongs("ID_ARTIST")%>"
                                  data-album="<%=rsSongs("ID_ALBUM")%>"
                                  onclick="handleEditSong(this)">
                            <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                            <span>Editar</span>
                          </button>

                          <!-- Botón Eliminar -->
                          <button type="button" class="btn-crud btn-crud-delete" title="Eliminar Canción"
                                  data-id="<%=songId%>"
                                  data-name="<%=SafeHtml(songName)%>"
                                  onclick="handleDeleteSong(this)">
                            <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                            <span>Eliminar</span>
                          </button>
                        </div>
                      </td>
                    <% End If %>
                  </tr>
                <%
                        rowIdx = rowIdx + 1
                        rsSongs.MoveNext
                    Loop
                Else
                    Dim colSpanCount
                    colSpanCount = 6
                    If isAdminUser Then colSpanCount = 7
                    Response.Write "<tr><td colspan=""" & colSpanCount & """ style=""text-align:center; padding: 36px; color: var(--cn-text-muted);"">No se encontraron pistas registradas en el catálogo.</td></tr>"
                End If
                CloseRecordset rsSongs
                %>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </main>
  </div>

  <!-- ====================================================================
       MODALES DE CRUD (Solo se renderizan si el usuario es Administrador)
       ==================================================================== -->
  <% If isAdminUser Then %>
    <!-- Modal 1: Agregar Canción (Admin) -->
    <div id="modalAddSong" class="modal-overlay">
      <div class="modal-content">
        <div class="modal-header">
          <h3 class="modal-title">Agregar Nueva Canción</h3>
          <button type="button" class="modal-close" onclick="closeModal('modalAddSong')">&times;</button>
        </div>
        <form action="canciones.asp" method="POST" autocomplete="off">
          <input type="hidden" name="action" value="create">

          <div class="form-group">
            <label class="form-label" for="addSongName">Título de la Canción *</label>
            <input type="text" id="addSongName" name="song_name" class="form-control" placeholder="Ej. Neon Nights" required maxlength="50">
          </div>

          <div class="form-group">
            <label class="form-label" for="addSongPath">Ruta del Archivo de Audio *</label>
            <select id="addSongPath" name="song_path" class="form-control" required>
              <option value="~/songFiles/bensound-adventure.mp3">~/songFiles/bensound-adventure.mp3 (Adventure)</option>
              <option value="~/songFiles/bensound-dreams.mp3">~/songFiles/bensound-dreams.mp3 (Dreams)</option>
              <option value="~/songFiles/bensound-energy.mp3">~/songFiles/bensound-energy.mp3 (Energy)</option>
              <option value="~/songFiles/bensound-scifi.mp3">~/songFiles/bensound-scifi.mp3 (Scifi)</option>
              <option value="~/songFiles/bensound-tomorrow.mp3">~/songFiles/bensound-tomorrow.mp3 (Tomorrow)</option>
            </select>
          </div>

          <div class="form-group">
            <label class="form-label" for="addArtistId">Artista Intérprete</label>
            <select id="addArtistId" name="artist_id" class="form-control">
              <option value="0">-- Sin artista asignado --</option>
              <%
              If Not rsArtists.EOF Then
                  Do While Not rsArtists.EOF
                      Response.Write "<option value=""" & rsArtists("ID_ARTIST") & """>" & SafeHtml(rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME")) & "</option>"
                      rsArtists.MoveNext
                  Loop
              End If
              %>
            </select>
          </div>

          <div class="form-group">
            <label class="form-label" for="addAlbumId">Álbum Discográfico</label>
            <select id="addAlbumId" name="album_id" class="form-control">
              <option value="0">-- Sencillo / Sin álbum --</option>
              <%
              If Not rsAlbums.EOF Then
                  Do While Not rsAlbums.EOF
                      Response.Write "<option value=""" & rsAlbums("ID_ALBUM") & """>" & SafeHtml(rsAlbums("TITLE_ALBUM")) & "</option>"
                      rsAlbums.MoveNext
                  Loop
              End If
              %>
            </select>
          </div>

          <div class="modal-actions">
            <button type="button" class="btn btn-secondary" onclick="closeModal('modalAddSong')">Cancelar</button>
            <button type="submit" class="btn btn-primary">Guardar Canción</button>
          </div>
        </form>
      </div>
    </div>

    <!-- Modal 2: Editar Canción (Admin) -->
    <div id="modalEditSong" class="modal-overlay">
      <div class="modal-content">
        <div class="modal-header">
          <h3 class="modal-title">Editar Canción</h3>
          <button type="button" class="modal-close" onclick="closeModal('modalEditSong')">&times;</button>
        </div>
        <form action="canciones.asp" method="POST" autocomplete="off">
          <input type="hidden" name="action" value="edit">
          <input type="hidden" id="editSongId" name="id_song" value="">

          <div class="form-group">
            <label class="form-label" for="editSongName">Título de la Canción *</label>
            <input type="text" id="editSongName" name="song_name" class="form-control" required maxlength="50">
          </div>

          <div class="form-group">
            <label class="form-label" for="editSongPath">Ruta del Archivo de Audio *</label>
            <input type="text" id="editSongPath" name="song_path" class="form-control" required maxlength="255">
          </div>

          <div class="form-group">
            <label class="form-label" for="editArtistId">Artista Intérprete</label>
            <select id="editArtistId" name="artist_id" class="form-control">
              <option value="0">-- Sin artista asignado --</option>
              <%
              If rsArtists.State = 1 Then rsArtists.MoveFirst
              Do While Not rsArtists.EOF
                  Response.Write "<option value=""" & rsArtists("ID_ARTIST") & """>" & SafeHtml(rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME")) & "</option>"
                  rsArtists.MoveNext
              Loop
              CloseRecordset rsArtists
              %>
            </select>
          </div>

          <div class="form-group">
            <label class="form-label" for="editAlbumId">Álbum Discográfico</label>
            <select id="editAlbumId" name="album_id" class="form-control">
              <option value="0">-- Sencillo / Sin álbum --</option>
              <%
              If rsAlbums.State = 1 Then rsAlbums.MoveFirst
              Do While Not rsAlbums.EOF
                  Response.Write "<option value=""" & rsAlbums("ID_ALBUM") & """>" & SafeHtml(rsAlbums("TITLE_ALBUM")) & "</option>"
                  rsAlbums.MoveNext
              Loop
              CloseRecordset rsAlbums
              %>
            </select>
          </div>

          <div class="modal-actions">
            <button type="button" class="btn btn-secondary" onclick="closeModal('modalEditSong')">Cancelar</button>
            <button type="submit" class="btn btn-primary">Actualizar Canción</button>
          </div>
        </form>
      </div>
    </div>
  <% End If %>

  <!-- Modal Global de Playlists & Scripts -->
  <script src="js/app.js"></script>
  <script>
    // --------------------------------------------------------------------------
    // Control de Reproducción Multimedia Integrada (<audio controls>)
    // Evita reproducción simultánea y actualiza contador de reproducciones
    // --------------------------------------------------------------------------
    document.addEventListener('DOMContentLoaded', function () {
      const audioPlayers = document.querySelectorAll('.native-audio-player');

      audioPlayers.forEach(player => {
        player.addEventListener('play', function () {
          // Pausar cualquier otro reproductor que esté activo
          audioPlayers.forEach(other => {
            if (other !== player && !other.paused) {
              other.pause();
            }
          });

          // Notificar incremento de reproducción al servidor
          const songId = player.getAttribute('data-song-id');
          if (songId) {
            fetch('play_counter.asp?id=' + encodeURIComponent(songId), { method: 'POST' })
              .then(res => res.json())
              .then(data => {
                if (data && data.success) {
                  const badge = document.getElementById('playsSpan-' + songId);
                  if (badge) {
                    const current = parseInt(badge.textContent.trim(), 10) || 0;
                    badge.textContent = current + 1;
                  }
                }
              })
              .catch(() => {});
          }
        });
      });
    });

    // --------------------------------------------------------------------------
    // Filtro en Vivo de la Tabla de Canciones
    // --------------------------------------------------------------------------
    function filterSongsTable() {
      const input = document.getElementById('songFilterInput');
      const filter = input.value.toLowerCase().trim();
      const rows = document.querySelectorAll('.song-table-row');

      rows.forEach(row => {
        const text = row.innerText.toLowerCase();
        if (text.includes(filter)) {
          row.style.display = '';
        } else {
          row.style.display = 'none';
        }
      });
    }

    // --------------------------------------------------------------------------
    // Modales de Administración (Agregar, Editar, Eliminar)
    // --------------------------------------------------------------------------
    function openAddSongModal() {
      openModal('modalAddSong');
    }

    function handleEditSong(btn) {
      document.getElementById('editSongId').value = btn.getAttribute('data-id');
      document.getElementById('editSongName').value = btn.getAttribute('data-name');
      document.getElementById('editSongPath').value = btn.getAttribute('data-path');
      document.getElementById('editArtistId').value = btn.getAttribute('data-artist') || '0';
      document.getElementById('editAlbumId').value = btn.getAttribute('data-album') || '0';
      openModal('modalEditSong');
    }

    function handleDeleteSong(btn) {
      var id = btn.getAttribute('data-id');
      var name = btn.getAttribute('data-name');
      if (confirm('¿Estás seguro de que deseas eliminar permanentemente la canción "' + name + '" del catálogo?')) {
        window.location.href = 'canciones.asp?action=delete&id=' + encodeURIComponent(id);
      }
    }

    function openEditSongModal(id, name, path, artistId, albumId) {
      document.getElementById('editSongId').value = id;
      document.getElementById('editSongName').value = name;
      document.getElementById('editSongPath').value = path;
      document.getElementById('editArtistId').value = artistId || '0';
      document.getElementById('editAlbumId').value = albumId || '0';
      openModal('modalEditSong');
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
