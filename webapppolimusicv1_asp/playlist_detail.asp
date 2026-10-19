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

Dim conn, cmdPl, rsPl
Set conn = GetConnection()

' 1. Obtener datos de la playlist
Set cmdPl = Server.CreateObject("ADODB.Command")
cmdPl.ActiveConnection = conn
cmdPl.CommandType = adCmdText
cmdPl.CommandText = "SELECT p.ID_PLAYLIST, p.ID_USER, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME " & _
    "FROM TBL_PLAYLIST p " & _
    "INNER JOIN TBL_USER u ON p.ID_USER = u.ID_USER " & _
    "WHERE p.ID_PLAYLIST = ?"
cmdPl.Parameters.Append cmdPl.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
Set rsPl = cmdPl.Execute()

If rsPl.EOF Then
    CloseRecordset rsPl
    CloseConnection conn
    Response.Redirect "playlists.asp"
    Response.End
End If

Dim plOwnerId, plTitle, plDate, plType, plPhoto, plRawPhoto, plOwnerName, canManage
plOwnerId   = rsPl("ID_USER")
plTitle     = rsPl("TITLE")
plDate      = rsPl("CREATION_DATE")
plType      = rsPl("TYPE")
plRawPhoto  = rsPl("PHOTO")
plPhoto     = ResolveUrl(rsPl("PHOTO"))
plOwnerName = rsPl("USERNAME")
CloseRecordset rsPl

' 2. Validación de Control de Acceso (RBAC):
' Si es privada (TYPE = 1) y el usuario actual no es el dueño ni admin, denegar
If CInt(plType) = 1 And (CLng(plOwnerId) <> CurrentUserId() And Not IsAdmin()) Then
    CloseConnection conn
    Response.Redirect "playlists.asp?msg=private_forbidden"
    Response.End
End If

' Puede gestionar si es el dueño o es administrador
canManage = (CLng(plOwnerId) = CurrentUserId() Or IsAdmin())

' Mensajes de estado
Dim msgCode, infoMsg
msgCode = Request.QueryString("msg")
infoMsg = ""
If msgCode = "song_added" Then
    infoMsg = "Canción añadida a la playlist con éxito."
ElseIf msgCode = "song_removed" Then
    infoMsg = "Canción retirada de la playlist."
ElseIf msgCode = "updated" Then
    infoMsg = "Detalles de la playlist actualizados."
End If

' 3. Obtener canciones de esta playlist
Dim cmdSongs, rsSongs
Set cmdSongs = Server.CreateObject("ADODB.Command")
cmdSongs.ActiveConnection = conn
cmdSongs.CommandType = adCmdText
cmdSongs.CommandText = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
    " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME, " & _
    "(SELECT TOP 1 alb.TITLE_ALBUM " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_NAME, " & _
    "(SELECT TOP 1 alb.ALBUM_IMAGE_PATH " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE " & _
    "FROM TBL_PLAYLIST_SONG ps " & _
    "INNER JOIN TBL_SONG s ON ps.ID_SONG = s.ID_SONG " & _
    "WHERE ps.ID_PLAYLIST = ? " & _
    "ORDER BY s.SONG_NAME ASC"
cmdSongs.Parameters.Append cmdSongs.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
Set rsSongs = cmdSongs.Execute()

' 4. Si puede gestionar, obtener canciones disponibles que aún no estén en la playlist
Dim rsAvailableSongs
If canManage Then
    Dim cmdAvail
    Set cmdAvail = Server.CreateObject("ADODB.Command")
    cmdAvail.ActiveConnection = conn
    cmdAvail.CommandType = adCmdText
    cmdAvail.CommandText = "SELECT s.ID_SONG, s.SONG_NAME, " & _
        "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
        " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME " & _
        "FROM TBL_SONG s " & _
        "WHERE s.ID_SONG NOT IN (SELECT ID_SONG FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?) " & _
        "ORDER BY s.SONG_NAME ASC"
    cmdAvail.Parameters.Append cmdAvail.CreateParameter("@ID_PLAYLIST", adInteger, adParamInput, , CLng(playlistId))
    Set rsAvailableSongs = cmdAvail.Execute()
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title><%=SafeHtml(plTitle)%> • Playlist PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <% If infoMsg <> "" Then %>
          <div class="alert alert-success">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
            <span><%=SafeHtml(infoMsg)%></span>
          </div>
        <% End If %>

        <!-- Hero de Playlist -->
        <div style="display: flex; gap: 28px; align-items: flex-end; padding: 20px 0 30px; border-bottom: 1px solid var(--border-subtle); flex-wrap: wrap;">
          <img src="<%=plPhoto%>" alt="<%=SafeHtml(plTitle)%>" style="width: 190px; height: 190px; border-radius: var(--radius-md); object-fit: cover; box-shadow: var(--shadow-md);" onerror="this.src='img/default-cover.svg'">
          <div>
            <div style="display: flex; gap: 8px; align-items: center; margin-bottom: 8px;">
              <span class="role-tag role-tag-admin">Playlist <%=FormatPlaylistType(plType)%></span>
              <% If CLng(plOwnerId) = CurrentUserId() Then %>
                <span class="role-tag role-tag-user">Tu Creación</span>
              <% End If %>
            </div>
            <h1 style="font-size: 38px; font-weight: 800; color: #fff; margin: 4px 0 10px; letter-spacing: -0.02em;"><%=SafeHtml(plTitle)%></h1>
            <p style="color: var(--text-muted); font-size: 14px;">Creada por <strong><%=SafeHtml(plOwnerName)%></strong> el <%=FormatDateStr(plDate)%></p>
            
            <div style="margin-top: 18px; display: flex; gap: 12px; align-items: center; flex-wrap: wrap;">
              <button type="button" id="btnPlayPlaylist" class="btn btn-primary">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
                <span>Reproducir Todo</span>
              </button>

              <% If canManage Then %>
                <button type="button" class="btn btn-secondary btn-sm" onclick="openModal('modalAddSong')">
                  <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
                  <span>Añadir Canciones</span>
                </button>
                <button type="button" class="btn btn-secondary btn-sm" onclick="openModal('modalEditPlaylist')">
                  <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                  <span>Editar</span>
                </button>
                <button type="button" class="btn btn-danger btn-sm" onclick="confirmAction('¿Deseas eliminar permanentemente esta playlist?', 'playlist_delete.asp?id=<%=playlistId%>')">
                  <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                  <span>Eliminar</span>
                </button>
              <% End If %>

              <a href="playlists.asp" class="btn btn-secondary btn-sm">Volver</a>
            </div>
          </div>
        </div>

        <!-- Tabla de Canciones en la Playlist -->
        <div style="background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-subtle); overflow: hidden; margin-top: 20px;">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width: 40px; text-align: center;">#</th>
                <th>Título</th>
                <th>Artista</th>
                <th>Álbum</th>
                <th style="width: 120px; text-align: right;">Reproducciones</th>
                <th style="width: 100px; text-align: center;">Acciones</th>
              </tr>
            </thead>
            <tbody>
              <%
              Dim idx
              idx = 1
              If Not rsSongs.EOF Then
                  Do While Not rsSongs.EOF
                      Dim songCover, songAudio, artistName, albumName
                      songCover  = ResolveUrl(rsSongs("ALBUM_IMAGE"))
                      songAudio  = ResolveUrl(rsSongs("SONG_PATH"))
                      artistName = rsSongs("ARTIST_NAME")
                      If IsNull(artistName) Or Trim(artistName) = "" Then artistName = "Artista PoliMusic"
                      albumName  = rsSongs("ALBUM_NAME")
                      If IsNull(albumName) Or Trim(albumName) = "" Then albumName = "Sencillo"
              %>
                <tr class="song-row">
                  <td style="text-align: center; color: var(--text-dim); font-weight: 700;"><%=idx%></td>
                  <td>
                    <div class="song-cell-main">
                      <img src="<%=songCover%>" alt="Cover" class="song-thumb" onerror="this.src='img/default-cover.svg'">
                      <div class="song-meta">
                        <span class="song-title-text"><%=SafeHtml(rsSongs("SONG_NAME"))%></span>
                      </div>
                    </div>
                  </td>
                  <td><%=SafeHtml(artistName)%></td>
                  <td><%=SafeHtml(albumName)%></td>
                  <td style="text-align: right; font-variant-numeric: tabular-nums;"><%=rsSongs("PLAYS")%></td>
                  <td>
                    <div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
                      <button type="button" class="table-play-btn"
                              data-song-id="<%=rsSongs("ID_SONG")%>"
                              data-song-name="<%=SafeHtml(rsSongs("SONG_NAME"))%>"
                              data-artist-name="<%=SafeHtml(artistName)%>"
                              data-audio-src="<%=songAudio%>"
                              data-cover-src="<%=songCover%>"
                              title="Reproducir <%=SafeHtml(rsSongs("SONG_NAME"))%>">
                        <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                      </button>

                      <!-- Si es dueño o admin, botón para quitar canción de esta playlist -->
                      <% If canManage Then %>
                        <button type="button" class="btn-icon delete" title="Quitar de la Playlist" onclick="confirmAction('¿Quitar esta canción de la playlist?', 'playlist_song_remove.asp?playlist_id=<%=playlistId%>&song_id=<%=rsSongs("ID_SONG")%>')">
                          <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13H5v-2h14v2z"/></svg>
                        </button>
                      <% End If %>
                    </div>
                  </td>
                </tr>
              <%
                      idx = idx + 1
                      rsSongs.MoveNext
                  Loop
              Else
                  Response.Write "<tr><td colspan=""6"" style=""text-align:center; padding:28px; color:var(--text-muted);"">Esta playlist no tiene canciones aún. ¡Haz clic en 'Añadir Canciones' para empezar!</td></tr>"
              End If
              CloseRecordset rsSongs
              %>
            </tbody>
          </table>
        </div>
      </div>
    </main>
  </div>

  <!-- Modal para Agregar Canción a la Playlist (Solo si canManage) -->
  <% If canManage Then %>
  <div id="modalAddSong" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Añadir Canción a <%=SafeHtml(plTitle)%></h3>
        <button type="button" class="modal-close" onclick="closeModal('modalAddSong')">&times;</button>
      </div>
      <form action="playlist_song_add.asp" method="POST">
        <input type="hidden" name="playlist_id" value="<%=playlistId%>">
        <input type="hidden" name="return_url" value="playlist_detail.asp?id=<%=playlistId%>">

        <div class="form-group">
          <label class="form-label" for="selSong">Selecciona una Canción</label>
          <select id="selSong" name="song_id" class="form-control" required>
            <%
            If Not rsAvailableSongs.EOF Then
                Do While Not rsAvailableSongs.EOF
                    Dim avArtist
                    avArtist = rsAvailableSongs("ARTIST_NAME")
                    If IsNull(avArtist) Or Trim(avArtist) = "" Then avArtist = "Varios"
                    Response.Write "<option value=""" & rsAvailableSongs("ID_SONG") & """>" & SafeHtml(rsAvailableSongs("SONG_NAME")) & " — " & SafeHtml(avArtist) & "</option>"
                    rsAvailableSongs.MoveNext
                Loop
            Else
                Response.Write "<option disabled>Todas las canciones del catálogo ya están en esta lista</option>"
            End If
            CloseRecordset rsAvailableSongs
            %>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalAddSong')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Añadir a Lista</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Modal para Editar la Playlist -->
  <div id="modalEditPlaylist" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Editar Playlist</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalEditPlaylist')">&times;</button>
      </div>
      <form action="playlist_save.asp" method="POST">
        <input type="hidden" name="action" value="edit">
        <input type="hidden" name="id_playlist" value="<%=playlistId%>">

        <div class="form-group">
          <label class="form-label" for="editTitle">Nombre de la Playlist *</label>
          <input type="text" id="editTitle" name="title" class="form-control" value="<%=SafeHtml(plTitle)%>" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="editType">Privacidad</label>
          <select id="editType" name="type" class="form-control">
            <option value="0" <%=IIf(CInt(plType) = 0, "selected", "")%>>Pública (Visible para toda la comunidad)</option>
            <option value="1" <%=IIf(CInt(plType) = 1, "selected", "")%>>Privada (Solo tú puedes verla)</option>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="editPhoto">Carátula / Portada</label>
          <select id="editPhoto" name="photo" class="form-control">
            <option value="~/playlistImages/default.jpg" <%=IIf(plRawPhoto = "~/playlistImages/default.jpg", "selected", "")%>>Portada Por Defecto</option>
            <option value="~/playlistImages/playlist1.png" <%=IIf(plRawPhoto = "~/playlistImages/playlist1.png", "selected", "")%>>Playlist Neón 1 (playlist1.png)</option>
            <option value="~/playlistImages/playlist2.png" <%=IIf(plRawPhoto = "~/playlistImages/playlist2.png", "selected", "")%>>Playlist Vibras 2 (playlist2.png)</option>
            <option value="~/playlistImages/playlist3.png" <%=IIf(plRawPhoto = "~/playlistImages/playlist3.png", "selected", "")%>>Playlist Ritmo 3 (playlist3.png)</option>
            <option value="~/playlistImages/playlist4.png" <%=IIf(plRawPhoto = "~/playlistImages/playlist4.png", "selected", "")%>>Playlist Acústico 4 (playlist4.png)</option>
            <option value="~/playlistImages/playlist5.png" <%=IIf(plRawPhoto = "~/playlistImages/playlist5.png", "selected", "")%>>Playlist Estudio 5 (playlist5.png)</option>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalEditPlaylist')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Guardar Cambios</button>
        </div>
      </form>
    </div>
  </div>
  <% End If %>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
  <script>
    const btnPlayPlaylist = document.getElementById('btnPlayPlaylist');
    if (btnPlayPlaylist) {
      btnPlayPlaylist.addEventListener('click', function () {
        const triggers = document.querySelectorAll('.table-play-btn');
        if (triggers.length > 0) {
          triggers[0].click();
        }
      });
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
