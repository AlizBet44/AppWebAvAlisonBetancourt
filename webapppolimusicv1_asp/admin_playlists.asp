<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<!--#include file="includes/auth_check.asp"-->
<%
RequireAdmin()

Dim conn
Set conn = GetConnection()

Dim infoMsg, errorMsg
infoMsg = ""
errorMsg = ""

' OPERACIÓN: ELIMINAR CUALQUIER PLAYLIST (ADMIN)
Dim action, delId
action = Request("action")
delId  = Request("id")

If action = "delete" And IsNumeric(delId) Then
    On Error Resume Next
    Dim cmdDel
    Set cmdDel = Server.CreateObject("ADODB.Command")
    cmdDel.ActiveConnection = conn
    cmdDel.CommandType = adCmdText
    cmdDel.CommandText = "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST = ?; DELETE FROM TBL_PLAYLIST WHERE ID_PLAYLIST = ?;"
    cmdDel.Parameters.Append cmdDel.CreateParameter("@P1", adInteger, adParamInput, , CLng(delId))
    cmdDel.Parameters.Append cmdDel.CreateParameter("@P2", adInteger, adParamInput, , CLng(delId))
    cmdDel.Execute()

    infoMsg = "Playlist eliminada por el administrador."
    On Error GoTo 0
End If

' Listado general de todas las playlists
Dim rsPlaylists
Set rsPlaylists = conn.Execute("SELECT p.ID_PLAYLIST, p.TITLE, p.CREATION_DATE, p.TYPE, p.PHOTO, u.USERNAME, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST_SONG ps WHERE ps.ID_PLAYLIST = p.ID_PLAYLIST) AS TRACK_COUNT " & _
    "FROM TBL_PLAYLIST p " & _
    "INNER JOIN TBL_USER u ON p.ID_USER = u.ID_USER " & _
    "ORDER BY p.CREATION_DATE DESC")
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CRUD Playlists • PoliMusic Admin</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/player.css">
</head>
<body>
  <div class="app-layout">
    <!--#include file="includes/sidebar.asp"-->

    <main class="app-main">
      <!--#include file="includes/header.asp"-->

      <div class="main-content">
        <% If infoMsg <> "" Then %>
          <div class="alert alert-success">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
            <span><%=SafeHtml(infoMsg)%></span>
          </div>
        <% End If %>

        <div class="section-header">
          <div>
            <h1 class="section-title">Auditoría Global de Playlists</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Control y supervisión de todas las listas de reproducción de la plataforma</p>
          </div>
          <button type="button" class="btn btn-primary" onclick="openModal('modalCreatePlaylist')">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Crear Playlist</span>
          </button>
        </div>

        <div class="admin-card">
          <div class="data-table-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 50px;">ID</th>
                  <th style="width: 60px;">Portada</th>
                  <th>Título</th>
                  <th>Propietario</th>
                  <th>Fecha Creación</th>
                  <th>Privacidad</th>
                  <th>Total Pistas</th>
                  <th style="width: 120px; text-align: center;">Acciones</th>
                </tr>
              </thead>
              <tbody>
                <%
                If Not rsPlaylists.EOF Then
                    Do While Not rsPlaylists.EOF
                        Dim plImg, typeBadge
                        plImg = ResolveUrl(rsPlaylists("PHOTO"))
                        If CInt(rsPlaylists("TYPE")) = 1 Then
                            typeBadge = "<span class=""role-tag role-tag-admin"">Privada</span>"
                        Else
                            typeBadge = "<span class=""role-tag role-tag-user"">Pública</span>"
                        End If
                %>
                  <tr>
                    <td><strong>#<%=rsPlaylists("ID_PLAYLIST")%></strong></td>
                    <td>
                      <img src="<%=plImg%>" alt="Cover" style="width: 40px; height: 40px; border-radius: 4px; object-fit: cover;" onerror="this.src='img/default-cover.svg'">
                    </td>
                    <td style="color:#fff; font-weight: 600;">
                      <a href="playlist_detail.asp?id=<%=rsPlaylists("ID_PLAYLIST")%>" style="color: #60a5fa;">
                        <%=SafeHtml(rsPlaylists("TITLE"))%>
                      </a>
                    </td>
                    <td><%=SafeHtml(rsPlaylists("USERNAME"))%></td>
                    <td><%=FormatDateStr(rsPlaylists("CREATION_DATE"))%></td>
                    <td><%=typeBadge%></td>
                    <td><%=rsPlaylists("TRACK_COUNT")%> canciones</td>
                    <td>
                      <div class="actions-cell" style="justify-content: center;">
                        <a href="playlist_detail.asp?id=<%=rsPlaylists("ID_PLAYLIST")%>" class="btn-icon" title="Ver / Gestionar">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M12 4.5C7 4.5 2.73 7.61 1 12c1.73 4.39 6 7.5 11 7.5s9.27-3.11 11-7.5c-1.73-4.39-6-7.5-11-7.5zM12 17c-2.76 0-5-2.24-5-5s2.24-5 5-5 5 2.24 5 5-2.24 5-5 5zm0-8c-1.66 0-3 1.34-3 3s1.34 3 3 3 3-1.34 3-3-1.34-3-3-3z"/></svg>
                        </a>
                        <button type="button" class="btn-icon delete" title="Eliminar Playlist"
                                data-id="<%=rsPlaylists("ID_PLAYLIST")%>"
                                data-title="<%=SafeHtml(rsPlaylists("TITLE"))%>"
                                onclick="handleDeletePlaylist(this)">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                        </button>
                      </div>
                    </td>
                  </tr>
                <%
                        rsPlaylists.MoveNext
                    Loop
                Else
                    Response.Write "<tr><td colspan=""8"" style=""text-align:center; padding:24px;"">No hay listas de reproducción registradas en el sistema.</td></tr>"
                End If
                CloseRecordset rsPlaylists
                %>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </main>
  </div>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
  <script>
    function handleDeletePlaylist(btn) {
      var id = btn.getAttribute('data-id');
      var title = btn.getAttribute('data-title');
      confirmAction('¿Eliminar la playlist \'' + title + '\'?', 'admin_playlists.asp?action=delete&id=' + id);
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
