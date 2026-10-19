<%
' ==============================================================================
' PoliMusic - Sidebar Modular Navigation
' ==============================================================================
Dim currentScript
currentScript = LCase(Mid(Request.ServerVariables("SCRIPT_NAME"), InStrRev(Request.ServerVariables("SCRIPT_NAME"), "/") + 1))
%>
<aside class="app-sidebar">
  <!-- Brand Logo & Evidence -->
  <div class="sidebar-brand">
    <a href="index.asp" title="PoliMusic Home">
      <img src="img/logo.svg" alt="PoliMusic Logo" style="height: 38px;">
    </a>
    <div class="sidebar-evidence-tag" title="Usuario Conectado">
      <span class="evidence-dot"></span>
      <span><%=SafeHtml(CurrentUsername())%></span>
    </div>
  </div>

  <!-- Primary Navigation -->
  <nav class="sidebar-nav">
    <a href="index.asp" class="nav-link <%=IIf(currentScript = "index.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M10 20v-6h4v6h5v-8h3L12 3 2 12h3v8z"/></svg>
      <span>Inicio</span>
    </a>
    <a href="canciones.asp" class="nav-link <%=IIf(currentScript = "canciones.asp" Or currentScript = "songs.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M12 3v10.55c-.59-.34-1.27-.55-2-.55-2.21 0-4 1.79-4 4s1.79 4 4 4 4-1.79 4-4V7h4V3h-6z"/></svg>
      <span>Canciones</span>
    </a>
    <a href="albums.asp" class="nav-link <%=IIf(currentScript = "albums.asp" Or currentScript = "album_detail.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 14.5c-2.49 0-4.5-2.01-4.5-4.5S9.51 7.5 12 7.5s4.5 2.01 4.5 4.5-2.01 4.5-4.5 4.5zm0-5.5c-.55 0-1 .45-1 1s.45 1 1 1 1-.45 1-1-.45-1-1-1z"/></svg>
      <span>Álbumes</span>
    </a>
    <a href="artists.asp" class="nav-link <%=IIf(currentScript = "artists.asp" Or currentScript = "artist_detail.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
      <span>Artistas</span>
    </a>
    <a href="playlists.asp" class="nav-link <%=IIf(currentScript = "playlists.asp" Or currentScript = "playlist_detail.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M15 6H3v2h12V6zm0 4H3v2h12v-2zM3 16h8v-2H3v2zM17 6v8.18c-.31-.11-.65-.18-1-.18-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3V8h3V6h-5z"/></svg>
      <span>Mis Playlists</span>
    </a>
  </nav>

  <!-- Admin Navigation (RBAC: Solo para userAdmin con UserType = 1) -->
  <% If IsAdmin() Then %>
  <div class="sidebar-heading">Administración</div>
  <nav class="sidebar-nav">
    <a href="admin_dashboard.asp" class="nav-link <%=IIf(currentScript = "admin_dashboard.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M3 13h8V3H3v10zm0 8h8v-6H3v6zm10 0h8V11h-8v10zm0-18v6h8V3h-8z"/></svg>
      <span>Panel Control</span>
    </a>
    <a href="admin_songs.asp" class="nav-link <%=IIf(currentScript = "admin_songs.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M12 3v9.28c-.47-.17-.97-.28-1.5-.28C8.01 12 6 14.01 6 16.5S8.01 21 10.5 21c2.31 0 4.2-1.75 4.45-4H15V6h4V3h-7z"/></svg>
      <span>CRUD Canciones</span>
    </a>
    <a href="admin_albums.asp" class="nav-link <%=IIf(currentScript = "admin_albums.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 17.93c-3.95-.49-7-3.85-7-7.93 0-.62.08-1.21.21-1.79L9 15v1c0 1.1.9 2 2 2v1.93zm6.9-2.54c-.26-.81-1-1.39-1.9-1.39h-1v-3c0-.55-.45-1-1-1H8v-2h2c.55 0 1-.45 1-1V7h2c1.1 0 2-.9 2-2v-.41c2.93 1.19 5 4.06 5 7.41 0 2.08-.8 3.97-2.1 5.39z"/></svg>
      <span>CRUD Álbumes</span>
    </a>
    <a href="admin_artists.asp" class="nav-link <%=IIf(currentScript = "admin_artists.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M16 11c1.66 0 2.99-1.34 2.99-3S17.66 5 16 5c-1.66 0-3 1.34-3 3s1.34 3 3 3zm-8 0c1.66 0 2.99-1.34 2.99-3S9.66 5 8 5C6.34 5 5 6.34 5 8s1.34 3 3 3zm0 2c-2.33 0-7 1.17-7 3.5V19h14v-2.5c0-2.33-4.67-3.5-7-3.5zm8 0c-.29 0-.62.02-.97.05 1.16.84 1.97 1.97 1.97 3.45V19h6v-2.5c0-2.33-4.67-3.5-7-3.5z"/></svg>
      <span>CRUD Artistas</span>
    </a>
    <a href="admin_users.asp" class="nav-link <%=IIf(currentScript = "admin_users.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
      <span>CRUD Usuarios</span>
    </a>
    <a href="admin_playlists.asp" class="nav-link <%=IIf(currentScript = "admin_playlists.asp", "active", "")%>">
      <svg viewBox="0 0 24 24"><path d="M4 6H2v14c0 1.1.9 2 2 2h14v-2H4V6zm16-4H8c-1.1 0-2 .9-2 2v12c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V4c0-1.1-.9-2-2-2zm-1 9h-4v4h-2v-4H9V9h4V5h2v4h4v2z"/></svg>
      <span>Todas Playlists</span>
    </a>
  </nav>
  <% End If %>

  <!-- Playlists del Usuario (Acceso Rápido) -->
  <div class="sidebar-heading">Tus Listas</div>
  <div class="sidebar-playlists">
    <%
    ' Consulta parametrizada de las playlists del usuario logueado
    On Error Resume Next
    Dim connNavPl, cmdNavPl, rsNavPl
    Set connNavPl = GetConnection()
    Set cmdNavPl = Server.CreateObject("ADODB.Command")
    cmdNavPl.ActiveConnection = connNavPl
    cmdNavPl.CommandType = adCmdText
    cmdNavPl.CommandText = "SELECT ID_PLAYLIST, TITLE FROM TBL_PLAYLIST WHERE ID_USER = ? ORDER BY CREATION_DATE DESC"
    cmdNavPl.Parameters.Append cmdNavPl.CreateParameter("@ID_USER", adInteger, adParamInput, , CurrentUserId())
    Set rsNavPl = cmdNavPl.Execute()

    If Not rsNavPl Is Nothing Then
      If Not rsNavPl.EOF Then
        Do While Not rsNavPl.EOF
          Response.Write "<a href=""playlist_detail.asp?id=" & rsNavPl("ID_PLAYLIST") & """ class=""playlist-link"">" & SafeHtml(rsNavPl("TITLE")) & "</a>"
          rsNavPl.MoveNext
        Loop
      Else
        Response.Write "<span style='font-size:12px; color:var(--text-dim); padding-left:12px;'>Sin listas aún</span>"
      End If
      CloseRecordset rsNavPl
    End If
    CloseConnection connNavPl
    On Error GoTo 0
    %>
  </div>

  <!-- Sidebar Footer: Usuario y Logout -->
  <div class="sidebar-footer">
    <div class="user-snippet">
      <img src="<%=CurrentUserPhoto()%>" alt="Foto de perfil" class="user-avatar-sm" onerror="this.src='UserImages/default.jpg'">
      <div class="user-info">
        <span class="user-name"><%=SafeHtml(CurrentUsername())%></span>
        <span class="user-role-badge"><%=CurrentUserRole()%></span>
      </div>
    </div>
    <a href="logout.asp" class="logout-btn-sidebar" title="Cerrar Sesión">
      <svg viewBox="0 0 24 24" width="20" height="20" fill="currentColor">
        <path d="M17 7l-1.41 1.41L18.17 11H8v2h10.17l-2.58 2.58L17 17l5-5zM4 5h8V3H4c-1.1 0-2 .9-2 2v14c0 1.1.9 2 2 2h8v-2H4V5z"/>
      </svg>
    </a>
  </div>
</aside>
<%
' Helper inline if function
Function IIf(expr, valTrue, valFalse)
    If expr Then IIf = valTrue Else IIf = valFalse
End Function
%>
