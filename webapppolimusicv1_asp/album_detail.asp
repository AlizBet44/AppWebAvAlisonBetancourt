<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim albumId
albumId = Request.QueryString("id")

If Not IsNumeric(albumId) Then
    Response.Redirect "albums.asp"
    Response.End
End If

Dim conn, cmdAlbum, rsAlbum
Set conn = GetConnection()

' 1. Obtener detalles del álbum con ADODB.Command parametrizado
Set cmdAlbum = Server.CreateObject("ADODB.Command")
cmdAlbum.ActiveConnection = conn
cmdAlbum.CommandType = adCmdText
cmdAlbum.CommandText = "SELECT ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM WHERE ID_ALBUM = ?"
cmdAlbum.Parameters.Append cmdAlbum.CreateParameter("@ID_ALBUM", adInteger, adParamInput, , CLng(albumId))
Set rsAlbum = cmdAlbum.Execute()

If rsAlbum.EOF Then
    CloseRecordset rsAlbum
    CloseConnection conn
    Response.Redirect "albums.asp"
    Response.End
End If

Dim albumTitle, albumYear, albumImg
albumTitle = rsAlbum("TITLE_ALBUM")
albumYear  = rsAlbum("RELEASE_YEAR")
albumImg   = ResolveUrl(rsAlbum("ALBUM_IMAGE_PATH"))
CloseRecordset rsAlbum

' 2. Obtener pistas del álbum
Dim cmdTracks, rsTracks
Set cmdTracks = Server.CreateObject("ADODB.Command")
cmdTracks.ActiveConnection = conn
cmdTracks.CommandType = adCmdText
cmdTracks.CommandText = "SELECT s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
    " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME " & _
    "FROM TBL_TRACKLIST tl " & _
    "INNER JOIN TBL_SONG s ON tl.ID_SONG = s.ID_SONG " & _
    "WHERE tl.ID_ALBUM = ? " & _
    "ORDER BY s.ID_SONG ASC"
cmdTracks.Parameters.Append cmdTracks.CreateParameter("@ID_ALBUM", adInteger, adParamInput, , CLng(albumId))
Set rsTracks = cmdTracks.Execute()

' Playlists del usuario para modal de agregar canción
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
  <title><%=SafeHtml(albumTitle)%> • Álbum PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <!-- Hero del Álbum -->
        <div style="display: flex; gap: 28px; align-items: flex-end; padding: 20px 0 30px; border-bottom: 1px solid var(--border-subtle); flex-wrap: wrap;">
          <img src="<%=albumImg%>" alt="<%=SafeHtml(albumTitle)%>" style="width: 190px; height: 190px; border-radius: var(--radius-md); object-fit: cover; box-shadow: var(--shadow-md);" onerror="this.src='img/default-cover.svg'">
          <div>
            <span class="role-tag role-tag-admin" style="margin-bottom: 8px;">Álbum Discográfico</span>
            <h1 style="font-size: 38px; font-weight: 800; color: #fff; margin: 4px 0 10px; letter-spacing: -0.02em;"><%=SafeHtml(albumTitle)%></h1>
            <p style="color: var(--text-muted); font-size: 14px;">Lanzamiento: <strong><%=SafeHtml(albumYear)%></strong> • PoliMusic Records</p>
            <div style="margin-top: 18px; display: flex; gap: 12px; align-items: center;">
              <button type="button" id="btnPlayWholeAlbum" class="btn btn-primary">
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
                <span>Reproducir Álbum</span>
              </button>
              <a href="albums.asp" class="btn btn-secondary btn-sm">Volver a Álbumes</a>
            </div>
          </div>
        </div>

        <!-- Lista de Pistas -->
        <div style="background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-subtle); overflow: hidden; margin-top: 20px;">
          <table class="song-table">
            <thead>
              <tr>
                <th style="width: 40px; text-align: center;">#</th>
                <th>Título</th>
                <th>Artista</th>
                <th style="width: 130px; text-align: right;">Reproducciones</th>
                <th style="width: 120px; text-align: center;">Acciones</th>
              </tr>
            </thead>
            <tbody>
              <%
              Dim trackIdx
              trackIdx = 1
              If Not rsTracks.EOF Then
                  Do While Not rsTracks.EOF
                      Dim songAudio, artistName
                      songAudio  = ResolveUrl(rsTracks("SONG_PATH"))
                      artistName = rsTracks("ARTIST_NAME")
                      If IsNull(artistName) Or Trim(artistName) = "" Then artistName = "Artista PoliMusic"
              %>
                <tr class="song-row">
                  <td style="text-align: center; color: var(--text-dim); font-weight: 700;"><%=trackIdx%></td>
                  <td>
                    <span class="song-title-text"><%=SafeHtml(rsTracks("SONG_NAME"))%></span>
                  </td>
                  <td><%=SafeHtml(artistName)%></td>
                  <td style="text-align: right; font-variant-numeric: tabular-nums;"><%=rsTracks("PLAYS")%></td>
                  <td>
                    <div style="display: flex; align-items: center; justify-content: center; gap: 8px;">
                      <button type="button" class="table-play-btn"
                              data-song-id="<%=rsTracks("ID_SONG")%>"
                              data-song-name="<%=SafeHtml(rsTracks("SONG_NAME"))%>"
                              data-artist-name="<%=SafeHtml(artistName)%>"
                              data-audio-src="<%=songAudio%>"
                              data-cover-src="<%=albumImg%>"
                              title="Reproducir <%=SafeHtml(rsTracks("SONG_NAME"))%>">
                        <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                      </button>

                      <% If plCount > 0 Then %>
                        <button type="button" class="btn-icon" title="Añadir a mi Playlist" onclick="openAddSongModal(<%=rsTracks("ID_SONG")%>, '<%=SafeHtml(Replace(rsTracks("SONG_NAME"), "'", "\'"))%>')">
                          <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M14 10H2v2h12v-2zm0-4H2v2h12V6zm4 8v-4h-2v4h-4v2h4v4h2v-4h4v-2h-4zM2 16h8v-2H2v2z"/></svg>
                        </button>
                      <% End If %>
                    </div>
                  </td>
                </tr>
              <%
                      trackIdx = trackIdx + 1
                      rsTracks.MoveNext
                  Loop
              Else
                  Response.Write "<tr><td colspan=""5"" style=""text-align:center; padding:24px;"">Este álbum no tiene pistas asociadas.</td></tr>"
              End If
              CloseRecordset rsTracks
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
        <input type="hidden" name="return_url" value="album_detail.asp?id=<%=albumId%>">
        
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

    // Botón para reproducir todo el álbum
    const btnPlayWholeAlbum = document.getElementById('btnPlayWholeAlbum');
    if (btnPlayWholeAlbum) {
      btnPlayWholeAlbum.addEventListener('click', function () {
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
