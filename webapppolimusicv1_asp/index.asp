<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
' Exigir autenticación obligatoria
RequireLogin()

Dim conn, cmd
Set conn = GetConnection()

' Datos para el Dashboard
Dim totalSongs, totalAlbums, totalArtists, totalUsers, totalPlaylists, totalPlays
totalSongs = 0
totalAlbums = 0
totalArtists = 0
totalUsers = 0
totalPlaylists = 0
totalPlays = 0

' Si es Administrador, obtener métricas del sistema
If IsAdmin() Then
    On Error Resume Next
    Dim rsStats
    Set rsStats = conn.Execute("SELECT " & _
        "(SELECT COUNT(*) FROM TBL_SONG) AS TotalSongs, " & _
        "(SELECT COUNT(*) FROM TBL_ALBUM) AS TotalAlbums, " & _
        "(SELECT COUNT(*) FROM TBL_ARTIST) AS TotalArtists, " & _
        "(SELECT COUNT(*) FROM TBL_USER) AS TotalUsers, " & _
        "(SELECT COUNT(*) FROM TBL_PLAYLIST) AS TotalPlaylists, " & _
        "(SELECT ISNULL(SUM(PLAYS), 0) FROM TBL_SONG) AS TotalPlays")
    If Not rsStats.EOF Then
        totalSongs     = rsStats("TotalSongs")
        totalAlbums    = rsStats("TotalAlbums")
        totalArtists   = rsStats("TotalArtists")
        totalUsers     = rsStats("TotalUsers")
        totalPlaylists = rsStats("TotalPlaylists")
        totalPlays     = rsStats("TotalPlays")
    End If
    CloseRecordset rsStats
    On Error GoTo 0
End If

' Saludo dinámico según la hora del día
Dim hourNow, greeting
hourNow = Hour(Now)
If hourNow < 12 Then
    greeting = "Buenos días"
ElseIf hourNow < 19 Then
    greeting = "Buenas tardes"
Else
    greeting = "Buenas noches"
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>PoliMusic • Tu Música EPN</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!-- Sidebar Modular -->
    <!--#include file="includes/sidebar.asp"-->

    <!-- Vista Principal -->
    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <!-- Banner Hero Dinámico -->
        <section class="page-hero">
          <div>
            <div class="hero-greeting"><%=greeting%>, <%=SafeHtml(CurrentUsername())%></div>
            <% If IsAdmin() Then %>
              <h1 class="hero-title">Centro de Control de PoliMusic</h1>
              <p class="hero-subtitle">Tienes permisos totales de Administrador. Administra canciones, artistas, álbumes, listas de reproducción y cuentas de usuario.</p>
            <% Else %>
              <h1 class="hero-title">Música en Vivo, Ritmo Politécnico</h1>
              <p class="hero-subtitle">Explora los éxitos más reproducidos, descubre nuevos artistas y personaliza tus propias playlists cuando quieras.</p>
            <% End If %>
          </div>

          <div style="z-index: 2; display: flex; gap: 10px;">
            <a href="canciones.asp" class="btn btn-primary">
              <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M8 5v14l11-7z"/></svg>
              <span>Explorar Todo</span>
            </a>
          </div>
        </section>

        <!-- DASHBOARD DINÁMICO: Si el usuario es Administrador (userAdmin) -->
        <% If IsAdmin() Then %>
          <section>
            <div class="section-header">
              <h2 class="section-title">Métricas Globales del Sistema</h2>
              <span class="role-tag role-tag-admin">Modo Auditoría</span>
            </div>
            <div class="stats-grid">
              <div class="stat-card">
                <div class="stat-icon">
                  <svg viewBox="0 0 24 24"><path d="M12 3v9.28c-.47-.17-.97-.28-1.5-.28C8.01 12 6 14.01 6 16.5S8.01 21 10.5 21c2.31 0 4.2-1.75 4.45-4H15V6h4V3h-7z"/></svg>
                </div>
                <span class="stat-label">Canciones</span>
                <span class="stat-number"><%=totalSongs%></span>
              </div>
              <div class="stat-card">
                <div class="stat-icon">
                  <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 14.5c-2.49 0-4.5-2.01-4.5-4.5S9.51 7.5 12 7.5s4.5 2.01 4.5 4.5-2.01 4.5-4.5 4.5zm0-5.5c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>
                </div>
                <span class="stat-label">Álbumes</span>
                <span class="stat-number"><%=totalAlbums%></span>
              </div>
              <div class="stat-card">
                <div class="stat-icon">
                  <svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>
                </div>
                <span class="stat-label">Artistas</span>
                <span class="stat-number"><%=totalArtists%></span>
              </div>
              <div class="stat-card">
                <div class="stat-icon">
                  <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
                </div>
                <span class="stat-label">Usuarios</span>
                <span class="stat-number"><%=totalUsers%></span>
              </div>
              <div class="stat-card">
                <div class="stat-icon">
                  <svg viewBox="0 0 24 24"><path d="M15 6H3v2h12V6zm0 4H3v2h12v-2zM3 16h8v-2H3v2zM17 6v8.18c-.31-.11-.65-.18-1-.18-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3V8h3V6h-5z"/></svg>
                </div>
                <span class="stat-label">Playlists</span>
                <span class="stat-number"><%=totalPlaylists%></span>
              </div>
              <div class="stat-card">
                <div class="stat-icon" style="color: #38bdf8;">
                  <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm-2 14.5v-9l6 4.5-6 4.5z"/></svg>
                </div>
                <span class="stat-label">Reproducciones</span>
                <span class="stat-number"><%=totalPlays%></span>
              </div>
            </div>
          </section>

          <!-- Accesos Rápidos de Administración -->
          <section>
            <div class="section-header">
              <h2 class="section-title">Accesos Rápidos de Mantenimiento</h2>
            </div>
            <div style="display: flex; gap: 12px; flex-wrap: wrap;">
              <a href="admin_songs.asp" class="btn btn-secondary">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M12 3v9.28c-.47-.17-.97-.28-1.5-.28C8.01 12 6 14.01 6 16.5S8.01 21 10.5 21c2.31 0 4.2-1.75 4.45-4H15V6h4V3h-7z"/></svg>
                <span>Gestionar Canciones</span>
              </a>
              <a href="admin_albums.asp" class="btn btn-secondary">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 14.5c-2.49 0-4.5-2.01-4.5-4.5S9.51 7.5 12 7.5s4.5 2.01 4.5 4.5-2.01 4.5-4.5 4.5zm0-5.5c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>
                <span>Gestionar Álbumes</span>
              </a>
              <a href="admin_artists.asp" class="btn btn-secondary">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>
                <span>Gestionar Artistas</span>
              </a>
              <a href="admin_users.asp" class="btn btn-secondary">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
                <span>Gestionar Usuarios</span>
              </a>
              <a href="admin_playlists.asp" class="btn btn-secondary">
                <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M15 6H3v2h12V6zm0 4H3v2h12v-2zM3 16h8v-2H3v2zM17 6v8.18c-.31-.11-.65-.18-1-.18-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3V8h3V6h-5z"/></svg>
                <span>Todas las Playlists</span>
              </a>
            </div>
          </section>
        <% End If %>

        <!-- SECCIÓN: Canciones Más Escuchadas (Top Plays) -->
        <section>
          <div class="section-header">
            <h2 class="section-title">Canciones en Tendencia</h2>
            <a href="canciones.asp" class="section-action">Ver todas</a>
          </div>

          <div style="background: var(--bg-card); border-radius: var(--radius-md); border: 1px solid var(--border-subtle); overflow: hidden;">
            <table class="song-table">
              <thead>
                <tr>
                  <th style="width: 40px; text-align: center;">#</th>
                  <th>Título</th>
                  <th>Álbum</th>
                  <th style="width: 140px; text-align: right;">Reproducciones</th>
                  <th style="width: 60px;"></th>
                </tr>
              </thead>
              <tbody>
                <%
                Dim rsTopSongs
                Set rsTopSongs = conn.Execute("SELECT TOP 5 " & _
                    "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
                    "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
                    " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME, " & _
                    "(SELECT TOP 1 alb.TITLE_ALBUM " & _
                    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_NAME, " & _
                    "(SELECT TOP 1 alb.ALBUM_IMAGE_PATH " & _
                    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_IMAGE " & _
                    "FROM TBL_SONG s ORDER BY ISNULL(s.PLAYS, 0) DESC, s.ID_SONG ASC")

                Dim idx
                idx = 1
                If Not rsTopSongs.EOF Then
                    Do While Not rsTopSongs.EOF
                        Dim songCover, songAudio, artistName, albumName
                        songCover  = ResolveUrl(rsTopSongs("ALBUM_IMAGE"))
                        songAudio  = ResolveUrl(rsTopSongs("SONG_PATH"))
                        artistName = rsTopSongs("ARTIST_NAME")
                        If IsNull(artistName) Or Trim(artistName) = "" Then artistName = "Artista PoliMusic"
                        albumName  = rsTopSongs("ALBUM_NAME")
                        If IsNull(albumName) Or Trim(albumName) = "" Then albumName = "Sencillo"
                %>
                  <tr class="song-row">
                    <td style="text-align: center; color: var(--text-dim); font-weight: 700;"><%=idx%></td>
                    <td>
                      <div class="song-cell-main">
                        <img src="<%=songCover%>" alt="Cover" class="song-thumb" onerror="this.src='img/default-cover.svg'">
                        <div class="song-meta">
                          <span class="song-title-text"><%=SafeHtml(rsTopSongs("SONG_NAME"))%></span>
                          <span class="song-artist-text"><%=SafeHtml(artistName)%></span>
                        </div>
                      </div>
                    </td>
                    <td><%=SafeHtml(albumName)%></td>
                    <td style="text-align: right; font-variant-numeric: tabular-nums;"><%=rsTopSongs("PLAYS")%></td>
                    <td>
                      <button type="button" class="table-play-btn" 
                              data-song-id="<%=rsTopSongs("ID_SONG")%>"
                              data-song-name="<%=SafeHtml(rsTopSongs("SONG_NAME"))%>"
                              data-artist-name="<%=SafeHtml(artistName)%>"
                              data-audio-src="<%=songAudio%>"
                              data-cover-src="<%=songCover%>"
                              title="Reproducir <%=SafeHtml(rsTopSongs("SONG_NAME"))%>">
                        <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                      </button>
                    </td>
                  </tr>
                <%
                        idx = idx + 1
                        rsTopSongs.MoveNext
                    Loop
                Else
                    Response.Write "<tr><td colspan=""5"" style=""text-align:center; padding:20px;"">No hay canciones registradas en el catálogo.</td></tr>"
                End If
                CloseRecordset rsTopSongs
                %>
              </tbody>
            </table>
          </div>
        </section>

        <!-- SECCIÓN: Álbumes Destacados -->
        <section>
          <div class="section-header">
            <h2 class="section-title">Álbumes Disponibles</h2>
            <a href="albums.asp" class="section-action">Explorar Álbumes</a>
          </div>
          <div class="grid-cards">
            <%
            Dim rsAlbums
            Set rsAlbums = conn.Execute("SELECT TOP 5 ID_ALBUM, TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH FROM TBL_ALBUM ORDER BY ID_ALBUM ASC")
            If Not rsAlbums.EOF Then
                Do While Not rsAlbums.EOF
                    Dim albImg
                    albImg = ResolveUrl(rsAlbums("ALBUM_IMAGE_PATH"))
            %>
              <div class="music-card" onclick="window.location.href='album_detail.asp?id=<%=rsAlbums("ID_ALBUM")%>'">
                <div class="card-cover-wrapper">
                  <img src="<%=albImg%>" alt="<%=SafeHtml(rsAlbums("TITLE_ALBUM"))%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
                  <button type="button" class="play-hover-btn" title="Ver Álbum">
                    <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                  </button>
                </div>
                <div class="card-title"><%=SafeHtml(rsAlbums("TITLE_ALBUM"))%></div>
                <div class="card-subtitle">Año <%=SafeHtml(rsAlbums("RELEASE_YEAR"))%> • Álbum</div>
              </div>
            <%
                    rsAlbums.MoveNext
                Loop
            End If
            CloseRecordset rsAlbums
            %>
          </div>
        </section>

        <!-- SECCIÓN: Artistas Populares -->
        <section>
          <div class="section-header">
            <h2 class="section-title">Artistas en PoliMusic</h2>
            <a href="artists.asp" class="section-action">Ver Artistas</a>
          </div>
          <div class="grid-cards">
            <%
            Dim rsArtists
            Set rsArtists = conn.Execute("SELECT TOP 5 ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE FROM TBL_ARTIST ORDER BY ID_ARTIST ASC")
            If Not rsArtists.EOF Then
                Do While Not rsArtists.EOF
                    Dim artImg, artFull
                    artImg = ResolveUrl(rsArtists("ARTIST_IMAGE"))
                    artFull = rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME")
            %>
              <div class="music-card artist-card" onclick="window.location.href='artist_detail.asp?id=<%=rsArtists("ID_ARTIST")%>'">
                <div class="card-cover-wrapper">
                  <img src="<%=artImg%>" alt="<%=SafeHtml(artFull)%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
                </div>
                <div class="card-title"><%=SafeHtml(artFull)%></div>
                <div class="card-subtitle">Artista</div>
              </div>
            <%
                    rsArtists.MoveNext
                Loop
            End If
            CloseRecordset rsArtists
            %>
          </div>
        </section>

        <!-- SECCIÓN: Playlists Públicas de la Comunidad -->
        <section>
          <div class="section-header">
            <h2 class="section-title">Playlists Públicas de la Comunidad EPN</h2>
            <a href="playlists.asp" class="section-action">Mis Playlists</a>
          </div>
          <div class="grid-cards">
            <%
            Dim rsPlPub
            Set rsPlPub = conn.Execute("SELECT TOP 5 p.ID_PLAYLIST, p.TITLE, p.PHOTO, u.USERNAME " & _
                "FROM TBL_PLAYLIST p INNER JOIN TBL_USER u ON p.ID_USER = u.ID_USER " & _
                "WHERE p.TYPE = 0 ORDER BY p.CREATION_DATE DESC")
            If Not rsPlPub.EOF Then
                Do While Not rsPlPub.EOF
                    Dim plImg
                    plImg = ResolveUrl(rsPlPub("PHOTO"))
            %>
              <div class="music-card" onclick="window.location.href='playlist_detail.asp?id=<%=rsPlPub("ID_PLAYLIST")%>'">
                <div class="card-cover-wrapper">
                  <img src="<%=plImg%>" alt="<%=SafeHtml(rsPlPub("TITLE"))%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
                  <button type="button" class="play-hover-btn" title="Abrir Playlist">
                    <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                  </button>
                </div>
                <div class="card-title"><%=SafeHtml(rsPlPub("TITLE"))%></div>
                <div class="card-subtitle">Por <%=SafeHtml(rsPlPub("USERNAME"))%></div>
              </div>
            <%
                    rsPlPub.MoveNext
                Loop
            End If
            CloseRecordset rsPlPub
            %>
          </div>
        </section>
      </div>
    </main>
  </div>

  <!-- Reproductor Persistente Spotify -->
  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
</body>
</html>
<%
CloseConnection conn
%>
