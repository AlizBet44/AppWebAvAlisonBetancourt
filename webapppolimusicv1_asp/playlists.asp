<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireLogin()

Dim conn
Set conn = GetConnection()

' Mensajes de estado
Dim msgCode, infoMsg, errorMsg
msgCode  = Request.QueryString("msg")
infoMsg  = ""
errorMsg = ""

If msgCode = "created" Then
    infoMsg = "¡Playlist creada exitosamente!"
ElseIf msgCode = "updated" Then
    infoMsg = "Playlist actualizada correctamente."
ElseIf msgCode = "deleted" Then
    infoMsg = "Playlist eliminada del sistema."
ElseIf msgCode = "private_forbidden" Then
    errorMsg = "No tienes permiso para acceder a una playlist privada de otro usuario."
ElseIf msgCode = "unauthorized_delete" Then
    errorMsg = "Solo puedes eliminar tus propias playlists."
End If

' 1. Obtener "Mis Playlists" (Creadas por el usuario autenticado)
Dim cmdMyPl, rsMyPl
Set cmdMyPl = Server.CreateObject("ADODB.Command")
cmdMyPl.ActiveConnection = conn
cmdMyPl.CommandType = adCmdText
cmdMyPl.CommandText = "SELECT p.ID_PLAYLIST, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TRACK_COUNT " & _
    "FROM TBL_PLAYLIST p " & _
    "WHERE p.ID_USER = ? " & _
    "ORDER BY p.CREATION_DATE DESC"
cmdMyPl.Parameters.Append cmdMyPl.CreateParameter("@ID_USER", adInteger, adParamInput, , CurrentUserId())
Set rsMyPl = cmdMyPl.Execute()

' 2. Obtener "Playlists Públicas de la Comunidad" (Otras personas con TYPE = 0)
Dim cmdPubPl, rsPubPl
Set cmdPubPl = Server.CreateObject("ADODB.Command")
cmdPubPl.ActiveConnection = conn
cmdPubPl.CommandType = adCmdText
cmdPubPl.CommandText = "SELECT p.ID_PLAYLIST, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TRACK_COUNT " & _
    "FROM TBL_PLAYLIST p " & _
    "INNER JOIN TBL_USER u ON p.ID_USER = u.ID_USER " & _
    "WHERE p.TYPE = 0 AND p.ID_USER <> ? " & _
    "ORDER BY p.CREATION_DATE DESC"
cmdPubPl.Parameters.Append cmdPubPl.CreateParameter("@ID_USER", adInteger, adParamInput, , CurrentUserId())
Set rsPubPl = cmdPubPl.Execute()
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Playlists • PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <!-- Alertas de estado -->
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

        <!-- SECCIÓN 1: Mis Playlists -->
        <div class="section-header">
          <div>
            <h1 class="section-title">Mis Playlists</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Listas creadas y administradas por ti. Puedes personalizarlas o agregar nuevas canciones.</p>
          </div>
          <button type="button" class="btn btn-primary btn-sm" onclick="openModal('modalCreatePlaylist')">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Crear Nueva Playlist</span>
          </button>
        </div>

        <div class="grid-cards" style="margin-bottom: 40px;">
          <%
          If Not rsMyPl.EOF Then
              Do While Not rsMyPl.EOF
                  Dim plImg, trackCount, typeLabel
                  plImg      = ResolveUrl(rsMyPl("PHOTO"))
                  trackCount = rsMyPl("TRACK_COUNT")
                  If CInt(rsMyPl("TYPE")) = 1 Then
                      typeLabel = "Privada"
                  Else
                      typeLabel = "Pública"
                  End If
          %>
            <div class="music-card">
              <div class="card-cover-wrapper" onclick="window.location.href='playlist_detail.asp?id=<%=rsMyPl("ID_PLAYLIST")%>'">
                <img src="<%=plImg%>" alt="<%=SafeHtml(rsMyPl("TITLE"))%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
                <button type="button" class="play-hover-btn" title="Abrir Playlist">
                  <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                </button>
              </div>
              <div class="card-title" onclick="window.location.href='playlist_detail.asp?id=<%=rsMyPl("ID_PLAYLIST")%>'"><%=SafeHtml(rsMyPl("TITLE"))%></div>
              <div class="card-subtitle"><%=typeLabel%> • <%=trackCount%> <%=IIf(trackCount = 1, "canción", "canciones")%></div>
              
              <!-- Botones de administración de la propia playlist (CRUD usuario normal y admin) -->
              <div style="display: flex; gap: 8px; margin-top: 10px; padding-top: 8px; border-top: 1px solid rgba(255,255,255,0.06);">
                <a href="playlist_detail.asp?id=<%=rsMyPl("ID_PLAYLIST")%>" class="btn btn-secondary btn-sm" style="flex:1;">Ver / Editar</a>
                <button type="button" class="btn-icon delete" title="Eliminar Playlist" onclick="confirmAction('¿Deseas eliminar permanentemente tu playlist <%=SafeHtml(Replace(rsMyPl("TITLE"), "'", "\'"))%>?', 'playlist_delete.asp?id=<%=rsMyPl("ID_PLAYLIST")%>')">
                  <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                </button>
              </div>
            </div>
          <%
                  rsMyPl.MoveNext
              Loop
          Else
          %>
            <div style="grid-column: 1 / -1; background: var(--bg-card); padding: 32px; border-radius: var(--radius-md); text-align: center; border: 1px dashed var(--border-subtle);">
              <svg viewBox="0 0 24 24" width="48" height="48" fill="var(--text-dim)" style="margin: 0 auto 12px;"><path d="M15 6H3v2h12V6zm0 4H3v2h12v-2zM3 16h8v-2H3v2zM17 6v8.18c-.31-.11-.65-.18-1-.18-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3V8h3V6h-5z"/></svg>
              <h3 style="color:#fff; margin-bottom: 6px;">Aún no tienes playlists creadas</h3>
              <p style="color: var(--text-muted); font-size: 13.5px; margin-bottom: 16px;">Crea tu primera lista de reproducción personalizada para organizar tu música</p>
              <button type="button" class="btn btn-primary btn-sm" onclick="openModal('modalCreatePlaylist')">Crear mi primera Playlist</button>
            </div>
          <%
          End If
          CloseRecordset rsMyPl
          %>
        </div>

        <!-- SECCIÓN 2: Playlists Públicas de la Comunidad Politécnica -->
        <div class="section-header">
          <div>
            <h2 class="section-title">Playlists Públicas de la Comunidad</h2>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Descubre las listas compartidas por otros estudiantes y miembros de PoliMusic</p>
          </div>
        </div>

        <div class="grid-cards">
          <%
          If Not rsPubPl.EOF Then
              Do While Not rsPubPl.EOF
                  Dim pubPlImg, pubTrackCount
                  pubPlImg      = ResolveUrl(rsPubPl("PHOTO"))
                  pubTrackCount = rsPubPl("TRACK_COUNT")
          %>
            <div class="music-card" onclick="window.location.href='playlist_detail.asp?id=<%=rsPubPl("ID_PLAYLIST")%>'">
              <div class="card-cover-wrapper">
                <img src="<%=pubPlImg%>" alt="<%=SafeHtml(rsPubPl("TITLE"))%>" class="card-cover" onerror="this.src='img/default-cover.svg'">
                <button type="button" class="play-hover-btn" title="Abrir Playlist">
                  <svg viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
                </button>
              </div>
              <div class="card-title"><%=SafeHtml(rsPubPl("TITLE"))%></div>
              <div class="card-subtitle">De <%=SafeHtml(rsPubPl("USERNAME"))%> • <%=pubTrackCount%> <%=IIf(pubTrackCount = 1, "canción", "canciones")%></div>
            </div>
          <%
                  rsPubPl.MoveNext
              Loop
          Else
              Response.Write "<p style=""color:var(--text-muted);"">No hay otras playlists públicas disponibles por ahora.</p>"
          End If
          CloseRecordset rsPubPl
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
