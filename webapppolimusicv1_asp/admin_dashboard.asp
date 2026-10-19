<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
' Exigir rol de Administrador obligatorio (RBAC)
RequireAdmin()

Dim conn
Set conn = GetConnection()

' Obtener estadísticas de auditoría
Dim rsStats, totalSongs, totalAlbums, totalArtists, totalUsers, totalPlaylists, totalPlays
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
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Panel de Administración • PoliMusic</title>
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
            <h1 class="section-title">Panel de Control Administrativo</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Control de Acceso y Gestión Integral de Entidades de PoliMusic</p>
          </div>
          <span class="role-tag role-tag-admin">Super Administrador</span>
        </div>

        <!-- Métricas KPI -->
        <div class="stats-grid">
          <div class="stat-card" onclick="window.location.href='admin_songs.asp'" style="cursor: pointer;">
            <div class="stat-icon">
              <svg viewBox="0 0 24 24"><path d="M12 3v9.28c-.47-.17-.97-.28-1.5-.28C8.01 12 6 14.01 6 16.5S8.01 21 10.5 21c2.31 0 4.2-1.75 4.45-4H15V6h4V3h-7z"/></svg>
            </div>
            <span class="stat-label">Canciones</span>
            <span class="stat-number"><%=totalSongs%></span>
          </div>

          <div class="stat-card" onclick="window.location.href='admin_albums.asp'" style="cursor: pointer;">
            <div class="stat-icon">
              <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 14.5c-2.49 0-4.5-2.01-4.5-4.5S9.51 7.5 12 7.5s4.5 2.01 4.5 4.5-2.01 4.5-4.5 4.5zm0-5.5c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>
            </div>
            <span class="stat-label">Álbumes</span>
            <span class="stat-number"><%=totalAlbums%></span>
          </div>

          <div class="stat-card" onclick="window.location.href='admin_artists.asp'" style="cursor: pointer;">
            <div class="stat-icon">
              <svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>
            </div>
            <span class="stat-label">Artistas</span>
            <span class="stat-number"><%=totalArtists%></span>
          </div>

          <div class="stat-card" onclick="window.location.href='admin_users.asp'" style="cursor: pointer;">
            <div class="stat-icon">
              <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
            </div>
            <span class="stat-label">Usuarios</span>
            <span class="stat-number"><%=totalUsers%></span>
          </div>

          <div class="stat-card" onclick="window.location.href='admin_playlists.asp'" style="cursor: pointer;">
            <div class="stat-icon">
              <svg viewBox="0 0 24 24"><path d="M15 6H3v2h12V6zm0 4H3v2h12v-2zM3 16h8v-2H3v2zM17 6v8.18c-.31-.11-.65-.18-1-.18-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3V8h3V6h-5z"/></svg>
            </div>
            <span class="stat-label">Playlists Totales</span>
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

        <!-- Módulos de Mantenimiento CRUD -->
        <div class="section-header" style="margin-top: 20px;">
          <h2 class="section-title">Módulos de Mantenimiento CRUD</h2>
        </div>

        <div style="display: grid; grid-template-columns: repeat(auto-fit, minmax(280px, 1fr)); gap: 20px;">
          <!-- Card CRUD Canciones -->
          <div class="admin-card">
            <h3 style="color: #fff; font-size: 16px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px;">
              <span style="color: var(--accent-pink);">🎵</span> Gestión de Canciones
            </h3>
            <p style="color: var(--text-muted); font-size: 13px; margin-bottom: 16px;">
              Crear, actualizar metadatos, asociar intérpretes y álbumes, y eliminar pistas musicales.
            </p>
            <a href="admin_songs.asp" class="btn btn-primary btn-sm">Abrir CRUD Canciones</a>
          </div>

          <!-- Card CRUD Álbumes -->
          <div class="admin-card">
            <h3 style="color: #fff; font-size: 16px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px;">
              <span style="color: var(--accent-pink);">💿</span> Gestión de Álbumes
            </h3>
            <p style="color: var(--text-muted); font-size: 13px; margin-bottom: 16px;">
              Administrar lanzamientos discográficos, años de publicación y carátulas de colección.
            </p>
            <a href="admin_albums.asp" class="btn btn-primary btn-sm">Abrir CRUD Álbumes</a>
          </div>

          <!-- Card CRUD Artistas -->
          <div class="admin-card">
            <h3 style="color: #fff; font-size: 16px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px;">
              <span style="color: var(--accent-pink);">🎤</span> Gestión de Artistas
            </h3>
            <p style="color: var(--text-muted); font-size: 13px; margin-bottom: 16px;">
              Registrar nuevos cantantes e intérpretes, editar perfiles y actualizar fotografías.
            </p>
            <a href="admin_artists.asp" class="btn btn-primary btn-sm">Abrir CRUD Artistas</a>
          </div>

          <!-- Card CRUD Usuarios -->
          <div class="admin-card">
            <h3 style="color: #fff; font-size: 16px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px;">
              <span style="color: var(--accent-pink);">👥</span> Gestión de Usuarios
            </h3>
            <p style="color: var(--text-muted); font-size: 13px; margin-bottom: 16px;">
              Crear usuarios, modificar correos, contraseñas y asignar roles de Administrador o Normal.
            </p>
            <a href="admin_users.asp" class="btn btn-primary btn-sm">Abrir CRUD Usuarios</a>
          </div>

          <!-- Card CRUD Playlists -->
          <div class="admin-card">
            <h3 style="color: #fff; font-size: 16px; margin-bottom: 8px; display: flex; align-items: center; gap: 8px;">
              <span style="color: var(--accent-pink);">📑</span> Gestión de Playlists
            </h3>
            <p style="color: var(--text-muted); font-size: 13px; margin-bottom: 16px;">
              Auditar y gestionar todas las listas de reproducción del sistema (públicas y privadas).
            </p>
            <a href="admin_playlists.asp" class="btn btn-primary btn-sm">Abrir CRUD Playlists</a>
          </div>
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
