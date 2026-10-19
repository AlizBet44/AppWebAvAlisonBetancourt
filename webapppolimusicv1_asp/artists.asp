<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim conn
Set conn = GetConnection()

' Obtener todos los artistas con conteo de canciones
Dim rsArtists
Set rsArtists = conn.Execute("SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
    "FROM TBL_ARTIST a ORDER BY a.ARTIST_NAME ASC")
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Artistas • PoliMusic</title>
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
            <h1 class="section-title">Artistas Destacados</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Explora los talentos e intérpretes musicales de PoliMusic</p>
          </div>
          <% If IsAdmin() Then %>
            <a href="admin_artists.asp" class="btn btn-primary btn-sm">
              <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
              <span>Gestionar Artistas (Admin)</span>
            </a>
          <% End If %>
        </div>

        <div class="grid-cards">
          <%
          If Not rsArtists.EOF Then
              Do While Not rsArtists.EOF
                  Dim artImg, artFull, songCount
                  artImg    = ResolveUrl(rsArtists("ARTIST_IMAGE"))
                  artFull   = rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME")
                  songCount = rsArtists("SONG_COUNT")
          %>
            <div class="music-card artist-card" onclick="window.location.href='artist_detail.asp?id=<%=rsArtists("ID_ARTIST")%>'">
              <div class="card-cover-wrapper">
                <img src="<%=artImg%>" alt="<%=SafeHtml(artFull)%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
              </div>
              <div class="card-title"><%=SafeHtml(artFull)%></div>
              <div class="card-subtitle"><%=songCount%> <%=IIf(songCount = 1, "canción disponible", "canciones disponibles")%></div>
            </div>
          <%
                  rsArtists.MoveNext
              Loop
          Else
              Response.Write "<p style=""color:var(--text-muted);"">No se encontraron artistas disponibles.</p>"
          End If
          CloseRecordset rsArtists
          %>
        </div>
      </div>
    </main>
  </div>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
</body>
</html>
<%
CloseConnection conn
%>
