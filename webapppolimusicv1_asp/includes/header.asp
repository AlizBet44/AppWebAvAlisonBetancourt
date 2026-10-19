<%
' ==============================================================================
' PoliMusic - Header Superior con Búsqueda y Perfil
' ==============================================================================
%>
<header class="app-header">
  <div class="header-left">
    <!-- Buscador en tiempo real de canciones, artistas y álbumes -->
    <div class="search-box">
      <svg viewBox="0 0 24 24"><path d="M15.5 14h-.79l-.28-.27C15.41 12.59 16 11.11 16 9.5 16 5.91 13.09 3 9.5 3S3 5.91 3 9.5 5.91 16 9.5 16c1.61 0 3.09-.59 4.23-1.57l.27.28v.79l5 4.99L20.49 19l-4.99-5zm-6 0C7.01 14 5 11.99 5 9.5S7.01 5 9.5 5 14 7.01 14 9.5 11.99 14 9.5 14z"/></svg>
      <input type="text" id="globalSearchInput" class="search-input" placeholder="Buscar canciones, artistas o álbumes...">
    </div>
  </div>

  <div class="header-right">
    <!-- Saludo Dinámico al Usuario Logueado -->
    <div class="user-greeting-badge" title="Sesión activa">
      <span class="user-greeting-dot"></span>
      <span class="user-greeting-prefix">Bienvenida,</span>
      <span class="user-greeting-name"><%=SafeHtml(CurrentUsername())%></span>
    </div>

    <!-- Botón Crear Playlist (Disponible para todos los usuarios) -->
    <button type="button" class="btn btn-secondary btn-sm" onclick="openModal('modalCreatePlaylist')">
      <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
      <span>Crear Playlist</span>
    </button>

    <!-- Indicador de Rol RBAC -->
    <% If IsAdmin() Then %>
      <span class="role-tag role-tag-admin" title="Privilegios de Administrador del Sistema">
        <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 1L3 5v6c0 5.55 3.84 10.74 9 12 5.16-1.26 9-6.45 9-12V5l-9-4zm0 10.99h7c-.53 4.12-3.28 7.79-7 8.94V12H5V6.3l7-3.11v8.8z"/></svg>
        Admin
      </span>
    <% Else %>
      <span class="role-tag role-tag-user" title="Perfil de Usuario EPN">
        <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 3c1.66 0 3 1.34 3 3s-1.34 3-3 3-3-1.34-3-3 1.34-3 3-3zm0 14.2c-2.5 0-4.71-1.28-6-3.22.03-1.99 4-3.08 6-3.08 1.99 0 5.97 1.09 6 3.08-1.29 1.94-3.5 3.22-6 3.22z"/></svg>
        EPN User
      </span>
    <% End If %>

    <a href="logout.asp" class="btn btn-danger btn-sm" title="Salir de PoliMusic">
      <span>Cerrar Sesión</span>
    </a>
  </div>
</header>

<!-- Modal Global para Crear Playlist Rápida -->
<div id="modalCreatePlaylist" class="modal-overlay">
  <div class="modal-content">
    <div class="modal-header">
      <h3 class="modal-title">Nueva Playlist</h3>
      <button type="button" class="modal-close" onclick="closeModal('modalCreatePlaylist')">&times;</button>
    </div>
    <form action="playlist_save.asp" method="POST">
      <input type="hidden" name="action" value="create">
      <div class="form-group">
        <label class="form-label" for="plTitle">Nombre de la Playlist *</label>
        <input type="text" id="plTitle" name="title" class="form-control" placeholder="Mi Playlist Favorita" required maxlength="50">
      </div>
      <div class="form-group">
        <label class="form-label" for="plType">Privacidad</label>
        <select id="plType" name="type" class="form-control">
          <option value="0">Pública (Visible para toda la comunidad PoliMusic)</option>
          <option value="1">Privada (Solo tú puedes verla)</option>
        </select>
      </div>
      <div class="form-group">
        <label class="form-label" for="plPhoto">Carátula / Portada</label>
        <select id="plPhoto" name="photo" class="form-control">
          <option value="~/playlistImages/default.jpg">Portada Por Defecto</option>
          <option value="~/playlistImages/playlist1.png">Playlist Neón 1</option>
          <option value="~/playlistImages/playlist2.png">Playlist Vibras 2</option>
          <option value="~/playlistImages/playlist3.png">Playlist Ritmo 3</option>
          <option value="~/playlistImages/playlist4.png">Playlist Acústico 4</option>
          <option value="~/playlistImages/playlist5.png">Playlist Estudio 5</option>
        </select>
      </div>
      <div class="modal-actions">
        <button type="button" class="btn btn-secondary" onclick="closeModal('modalCreatePlaylist')">Cancelar</button>
        <button type="submit" class="btn btn-primary">Crear Playlist</button>
      </div>
    </form>
  </div>
</div>
