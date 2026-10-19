<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim artistId
artistId = Request.QueryString("id")

If Not IsNumeric(artistId) Then
    Response.Redirect "artists.asp"
    Response.End
End If

Dim conn, cmdArtist, rsArtist
Set conn = GetConnection()

' 1. Obtener detalles del artista de forma parametrizada
Set cmdArtist = Server.CreateObject("ADODB.Command")
cmdArtist.ActiveConnection = conn
cmdArtist.CommandType = adCmdText
cmdArtist.CommandText = "SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST WHERE ID_ARTIST = ?"
cmdArtist.Parameters.Append cmdArtist.CreateParameter("@ID_ARTIST", adInteger, adParamInput, , CLng(artistId))
Set rsArtist = cmdArtist.Execute()

If rsArtist.EOF Then
    CloseRecordset rsArtist
    CloseConnection conn
    Response.Redirect "artists.asp"
    Response.End
End If

Dim artistFullName, artistImg
artistFullName = rsArtist("ARTIST_NAME") & " " & rsArtist("ARTIST_LASTNAME")
artistImg      = ResolveUrl(rsArtist("ARTIST_IMAGE"))
CloseRecordset rsArtist

' 2. Obtener canciones interpretadas por este artista
Dim cmdSongs, rsSongs
Set cmdSongs = Server.CreateObject("ADODB.Command")
cmdSongs.ActiveConnection = conn
cmdSongs.CommandType = adCmdText
cmdSongs.CommandText = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 alb.TITLE_ALBUM FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_NAME, " & _
    "(SELECT TOP 1 alb.ALBUM_IMAGE_PATH FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE " & _
    "FROM TBL_PLAYER p " & _
    "INNER JOIN TBL_SONG s ON p.ID_SONG = s.ID_SONG " & _
    "WHERE p.ID_ARTIST = ? " & _
    "ORDER BY s.SONG_NAME ASC"
cmdSongs.Parameters.Append cmdSongs.CreateParameter("@ID_ARTIST", adInteger, adParamInput, , CLng(artistId))
Set rsSongs = cmdSongs.Execute()

' Playlists del usuario
Dim cmdUserPl, rsUserPl, userPlaylists(), plCount, i
Set cmdUserPl = Server.CreateObject("ADODB.Command")
cmdUserPl.ActiveConnection = conn
cmdUserPl.CommandType = adCmdText
cmdUserPl.CommandText = "SELECT ID_PLAYLIST, TITLE FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY TITLE ASC"
cmdUserPl.Parameters.Append cmdUserPl.CreateParameter("@ID_USER", adInteger, adParamInput, , CurrentUserId())
Set rsUserPl = cmdUserPl.Execute()

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
  <title><%=SafeHtml(artistFullName)%> • Artista PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <!-- Hero del Artista -->
        <div style="display: flex; gap: 28px; align-items: center; padding: 20px 0 30px; border-bottom: 1px solid var(--border-subtle); flex-wrap: wrap;">
          <img src="<%=artistImg%>" alt="<%=SafeHtml(artistFullName)%>" style="width: 170px; height: 170px; border-radius: var(--radius-full); object-fit: cover; box-shadow: var(--shadow-md); border: 3px solid var(--primary);" onerror="this.src='img/default-cover.svg'">
          <div>
            <span class="role-tag role-tag-admin" style="margin-bottom: 8px;">Artista Verificado</span>
            <h1 style="font-size: 38px; font-weight: 800; color: #fff; margin: 4px 0 10px; letter-spacing: -0.02em;"><%=SafeHtml(artistFullName)%></h1>
            <p style="color: var(--text-muted); font-size: 14px;">Intérprete destacado en la plataforma PoliMusic</p>
            <div style="margin-top: 18px; display: flex; gap: 12px; align-items: center;">
              <button type="button" id="btnPlayAllArtist" class="btn btn-primary">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
                <span>Reproducir Canciones</span>
              </button>
              <a href="artists.asp" class="btn btn-secondary btn-sm">Volver a Artistas</a>
            </div>
          </div>
        </div>

        <!-- Repertorio de Canciones -->
        <div style="background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-subtle); overflow: hidden; margin-top: 20px;">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width: 40px; text-align: center;">#</th>
                <th>Título</th>
                <th>Álbum</th>
                <th style="width: 130px; text-align: right;">Reproducciones</th>
                <th style="width: 120px; text-align: center;">Acciones</th>
              </tr>
            </thead>
            <tbody>
              <%
              Dim songIdx
              songIdx = 1
              If Not rsSongs.EOF Then
                  Do While Not rsSongs.EOF
                      Dim songCover, songAudio, albumName
                      songCover = ResolveUrl(rsSongs("ALBUM_IMAGE"))
                      songAudio = ResolveUrl(rsSongs("SONG_PATH"))
                      albumName = rsSongs("ALBUM_NAME")
                      If IsNull(albumName) Or Trim(albumName) = "" Then albumName = "Sencillo"
              %>
                <tr class="song-row">
                  <td style="text-align: center; color: var(--text-dim); font-weight: 700;"><%=songIdx%></td>
                  <td>
                    <div class="song-cell-main">
                      <img src="<%=songCover%>" alt="Cover" class="song-thumb" onerror="this.src='img/default-cover.svg'">
                      <div class="song-meta">
                        <span class="song-title-text"><%=SafeHtml(rsSongs("SONG_NAME"))%></span>
                        <span class="song-artist-text"><%=SafeHtml(artistFullName)%></span>
                      </div>
                    </div>
                  </td>
                  <td><%=SafeHtml(albumName)%></td>
                  <td style="text-align: right; font-variant-numeric: tabular-nums;"><%=rsSongs("PLAYS")%></td>
                  <td>
                    <div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
                      <button type="button" class="table-play-btn"
                              data-song-id="<%=rsSongs("ID_SONG")%>"
                              data-song-name="<%=SafeHtml(rsSongs("SONG_NAME"))%>"
                              data-artist-name="<%=SafeHtml(artistFullName)%>"
                              data-audio-src="<%=songAudio%>"
                              data-cover-src="<%=songCover%>"
                              title="Reproducir <%=SafeHtml(rsSongs("SONG_NAME"))%>">
                        <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                      </button>

                      <% If plCount > 0 Then %>
                        <button type="button" class="btn-icon" title="Añadir a mi Playlist" onclick="openAddSongModal(<%=rsSongs("ID_SONG")%>, '<%=SafeHtml(Replace(rsSongs("SONG_NAME"), "'", "\'"))%>')">
                          <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M14 10H2v2h12v-2zm0-4H2v2h12V6zm4 8v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zM2 16h8v-2H2v2z"/></svg>
                        </button>
                      <% End If %>
                    </div>
                  </td>
                </tr>
              <%
                      songIdx = songIdx + 1
                      rsSongs.MoveNext
                  Loop
              Else
                  Response.Write "<tr><td colspan=""5"" style=""text-align:center; padding:24px;"">Este artista no tiene canciones asociadas en el sistema.</td></tr>"
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
        <input type="hidden" name="return_url" value="artist_detail.asp?id=<%=artistId%>">
        
        <p style="margin-bottom: 14px; font-size: 13.5px; color: var(--text-muted);">
          Selecciona tu playlist para agregar <strong id="modalAddSongTitle" style="color:#fff;"></strong>:
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

    const btnPlayAllArtist = document.getElementById('btnPlayAllArtist');
    if (btnPlayAllArtist) {
      btnPlayAllArtist.addEventListener('click', function () {
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
