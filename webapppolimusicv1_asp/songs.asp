<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()
Response.Redirect "canciones.asp"
Response.End
%>

' Consulta de canciones con artista y álbum
Dim rsSongs
Set rsSongs = conn.Execute("SELECT " & _
    "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
    " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME, " & _
    "(SELECT TOP 1 alb.TITLE_ALBUM " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_NAME, " & _
    "(SELECT TOP 1 alb.ALBUM_IMAGE_PATH " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE " & _
    "FROM TBL_SONG s ORDER BY s.SONG_NAME ASC")

' Obtener las playlists del usuario para el selector "Añadir a Playlist"
Dim cmdUserPl, rsUserPl
Set cmdUserPl = Server.CreateObject("ADODB.Command")
cmdUserPl.ActiveConnection = conn
cmdUserPl.CommandType = adCmdText
cmdUserPl.CommandText = "SELECT ID_PLAYLIST, TITLE FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY TITLE ASC"
cmdUserPl.Parameters.Append cmdUserPl.CreateParameter("@ID_USER", adInteger, adParamInput, , CurrentUserId())
Set rsUserPl = cmdUserPl.Execute()

' Cargar playlists en un arreglo para reutilizar en cada fila
Dim userPlaylists()
Dim plCount, i
plCount = 0
If Not rsUserPl.EOF Then
    Do While Not rsUserPl.EOF
        ReDim Preserve userPlaylists(1, plCount)
        userPlaylists(0, plCount) = rsUserPl("ID_PLAYLIST")
        userPlaylists(1, plCount) = rsUserPl("TITLE")
        plCount = plCount + 1
        rsUserPl.MoveNext
    Loop
End If
CloseRecordset rsUserPl
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Explorar Canciones • PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <div class="section-header">
          <div>
            <h1 class="section-title">Todas las Canciones</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Explora el catálogo completo de PoliMusic y escucha tus canciones preferidas</p>
          </div>
          <% If IsAdmin() Then %>
            <a href="admin_songs.asp" class="btn btn-primary btn-sm">
              <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
              <span>Añadir Nueva Canción (Admin)</span>
            </a>
          <% End If %>
        </div>

        <div style="background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-subtle); overflow: hidden;">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width: 40px; text-align: center;">#</th>
                <th>Título</th>
                <th>Artista</th>
                <th>Álbum</th>
                <th style="width: 130px; text-align: right;">Reproducciones</th>
                <th style="width: 140px; text-align: center;">Acciones</th>
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
                      <!-- Botón de Reproducción -->
                      <button type="button" class="table-play-btn"
                              data-song-id="<%=rsSongs("ID_SONG")%>"
                              data-song-name="<%=SafeHtml(rsSongs("SONG_NAME"))%>"
                              data-artist-name="<%=SafeHtml(artistName)%>"
                              data-audio-src="<%=songAudio%>"
                              data-cover-src="<%=songCover%>"
                              title="Reproducir <%=SafeHtml(rsSongs("SONG_NAME"))%>">
                        <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                      </button>

                      <!-- Botón Añadir a Playlist -->
                      <% If plCount > 0 Then %>
                        <button type="button" class="btn-icon" title="Añadir a mi Playlist" onclick="openAddSongModal(<%=rsSongs("ID_SONG")%>, '<%=SafeHtml(Replace(rsSongs("SONG_NAME"), "'", "\'"))%>')">
                          <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M14 10H2v2h12v-2zm0-4H2v2h12V6zm4 8v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zM2 16h8v-2H2v2z"/></svg>
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
                  Response.Write "<tr><td colspan=""6"" style=""text-align:center; padding:24px;"">No se encontraron canciones en el catálogo.</td></tr>"
              End If
              CloseRecordset rsSongs
              %>
            </tbody>
          </table>
        </div>
      </div>
    </main>
  </div>

  <!-- Modal para Agregar Canción a una Playlist -->
  <div id="modalAddToPlaylist" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Añadir a Playlist</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalAddToPlaylist')">&times;</button>
      </div>
      <form action="playlist_song_add.asp" method="POST">
        <input type="hidden" id="modalAddSongId" name="song_id" value="">
        <input type="hidden" name="return_url" value="songs.asp">
        
        <p style="margin-bottom: 14px; font-size: 13.5px; color: var(--text-muted);">
          Selecciona una de tus playlists para agregar <strong id="modalAddSongTitle" style="color:#fff;"></strong>:
        </p>

        <div class="form-group">
          <label class="form-label" for="selectPlaylist">Tus Playlists</label>
          <select id="selectPlaylist" name="playlist_id" class="form-control" required>
            <%
            For i = 0 To plCount - 1
                Response.Write "<option value=""" & userPlaylists(0, i) & """>" & SafeHtml(userPlaylists(1, i)) & "</option>"
            Next
            %>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalAddToPlaylist')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Añadir Canción</button>
        </div>
      </form>
    </div>
  </div>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
  <script>
    function openAddSongModal(songId, songTitle) {
      document.getElementById('modalAddSongId').value = songId;
      document.getElementById('modalAddSongTitle').textContent = songTitle;
      openModal('modalAddToPlaylist');
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
