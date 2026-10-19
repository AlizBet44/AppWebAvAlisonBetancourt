<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim conn
Set conn = GetConnection()

' Obtener todos los álbumes con conteo de pistas
Dim rsAlbums
Set rsAlbums = conn.Execute("SELECT a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = a.ID_ALBUM) AS TRACK_COUNT " & _
    "FROM TBL_ALBUM a ORDER BY a.TITLE_ALBUM ASC")
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Álbumes • PoliMusic</title>
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
            <h1 class="section-title">Discografía y Álbumes</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Explora colecciones completas y álbumes discográficos de PoliMusic</p>
          </div>
          <% If IsAdmin() Then %>
            <a href="admin_albums.asp" class="btn btn-primary btn-sm">
              <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
              <span>Gestionar Álbumes (Admin)</span>
            </a>
          <% End If %>
        </div>

        <div class="grid-cards">
          <%
          If Not rsAlbums.EOF Then
              Do While Not rsAlbums.EOF
                  Dim albImg, trackCount
                  albImg = ResolveUrl(rsAlbums("ALBUM_IMAGE_PATH"))
                  trackCount = rsAlbums("TRACK_COUNT")
          %>
            <div class="music-card" onclick="window.location.href='album_detail.asp?id=<%=rsAlbums("ID_ALBUM")%>'">
              <div class="card-cover-wrapper">
                <img src="<%=albImg%>" alt="<%=SafeHtml(rsAlbums("TITLE_ALBUM"))%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
                <button type="button" class="play-hover-btn" title="Ver Álbum">
                  <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                </button>
              </div>
              <div class="card-title"><%=SafeHtml(rsAlbums("TITLE_ALBUM"))%></div>
              <div class="card-subtitle"><%=SafeHtml(rsAlbums("RELEASE_YEAR"))%> • <%=trackCount%> <%=IIf(trackCount = 1, "canción", "canciones")%></div>
            </div>
          <%
                  rsAlbums.MoveNext
              Loop
          Else
              Response.Write "<p style=""color:var(--text-muted);"">No se encontraron álbumes disponibles.</p>"
          End If
          CloseRecordset rsAlbums
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
