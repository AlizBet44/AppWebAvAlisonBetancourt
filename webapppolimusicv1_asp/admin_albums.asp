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

Dim reqMethod, action
reqMethod = Request.ServerVariables("REQUEST_METHOD")
action    = Request("action")

' OPERACIÓN: CREAR ÁLBUM
If reqMethod = "POST" And action = "create" Then
    Dim newTitle, newYear, newImage
    newTitle = Trim(Request.Form("title_album"))
    newYear  = Request.Form("release_year")
    newImage = Trim(Request.Form("album_image_path"))

    If newTitle = "" Then
        errorMsg = "El título del álbum es obligatorio."
    Else
        If Not IsNumeric(newYear) Then newYear = Year(Now)
        If newImage = "" Then newImage = "~/albumImages/album1.png"

        On Error Resume Next
        Dim cmdInsAlb
        Set cmdInsAlb = Server.CreateObject("ADODB.Command")
        cmdInsAlb.ActiveConnection = conn
        cmdInsAlb.CommandType = adCmdText
        cmdInsAlb.CommandText = "INSERT INTO TBL_ALBUM (TITLE_ALBUM, RELEASE_YEAR, ALBUM_IMAGE_PATH) VALUES (?, ?, ?)"
        cmdInsAlb.Parameters.Append cmdInsAlb.CreateParameter("@TITLE", adVarChar, adParamInput, 50, newTitle)
        cmdInsAlb.Parameters.Append cmdInsAlb.CreateParameter("@YEAR", adInteger, adParamInput, , CInt(newYear))
        cmdInsAlb.Parameters.Append cmdInsAlb.CreateParameter("@IMG", adVarChar, adParamInput, 255, newImage)
        cmdInsAlb.Execute()

        If Err.Number <> 0 Then
            errorMsg = "Error al crear el álbum: " & Err.Description
            Err.Clear
        Else
            infoMsg = "Álbum '" & newTitle & "' registrado exitosamente."
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: EDITAR ÁLBUM
ElseIf reqMethod = "POST" And action = "edit" Then
    Dim editId, editTitle, editYear, editImage
    editId    = Request.Form("id_album")
    editTitle = Trim(Request.Form("title_album"))
    editYear  = Request.Form("release_year")
    editImage = Trim(Request.Form("album_image_path"))

    If Not IsNumeric(editId) Or editTitle = "" Then
        errorMsg = "Datos incompletos para actualizar el álbum."
    Else
        If Not IsNumeric(editYear) Then editYear = Year(Now)
        If editImage = "" Then editImage = "~/albumImages/album1.png"

        On Error Resume Next
        Dim cmdUpdAlb
        Set cmdUpdAlb = Server.CreateObject("ADODB.Command")
        cmdUpdAlb.ActiveConnection = conn
        cmdUpdAlb.CommandType = adCmdText
        cmdUpdAlb.CommandText = "UPDATE TBL_ALBUM SET TITLE_ALBUM = ?, RELEASE_YEAR = ?, ALBUM_IMAGE_PATH = ? WHERE ID_ALBUM = ?"
        cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@TITLE", adVarChar, adParamInput, 50, editTitle)
        cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@YEAR", adInteger, adParamInput, , CInt(editYear))
        cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@IMG", adVarChar, adParamInput, 255, editImage)
        cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@ID", adInteger, adParamInput, , CLng(editId))
        cmdUpdAlb.Execute()

        If Err.Number <> 0 Then
            errorMsg = "Error al actualizar el álbum: " & Err.Description
            Err.Clear
        Else
            infoMsg = "Álbum actualizado correctamente."
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: ELIMINAR ÁLBUM
ElseIf action = "delete" Then
    Dim delId
    delId = Request("id")
    If IsNumeric(delId) Then
        On Error Resume Next
        Dim cmdDelAlb
        Set cmdDelAlb = Server.CreateObject("ADODB.Command")
        cmdDelAlb.ActiveConnection = conn
        cmdDelAlb.CommandType = adCmdText
        cmdDelAlb.CommandText = "DELETE FROM TBL_TRACKLIST WHERE ID_ALBUM = ?; DELETE FROM TBL_ALBUM WHERE ID_ALBUM = ?;"
        cmdDelAlb.Parameters.Append cmdDelAlb.CreateParameter("@P1", adInteger, adParamInput, , CLng(delId))
        cmdDelAlb.Parameters.Append cmdDelAlb.CreateParameter("@P2", adInteger, adParamInput, , CLng(delId))
        cmdDelAlb.Execute()

        infoMsg = "Álbum eliminado correctamente."
        On Error GoTo 0
    End If
End If

' Listado de álbumes
Dim rsAlbums
Set rsAlbums = conn.Execute("SELECT a.ID_ALBUM, a.TITLE_ALBUM, a.RELEASE_YEAR, a.ALBUM_IMAGE_PATH, " & _
    "(SELECT COUNT(*) FROM TBL_TRACKLIST WHERE ID_ALBUM = a.ID_ALBUM) AS TRACK_COUNT " & _
    "FROM TBL_ALBUM a ORDER BY a.ID_ALBUM DESC")
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CRUD Álbumes • PoliMusic Admin</title>
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

        <% If errorMsg <> "" Then %>
          <div class="alert alert-danger">
            <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
            <span><%=SafeHtml(errorMsg)%></span>
          </div>
        <% End If %>

        <div class="section-header">
          <div>
            <h1 class="section-title">Gestión de Álbumes (CRUD)</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Crear, editar carátulas, años de lanzamiento y administrar discografía</p>
          </div>
          <button type="button" class="btn btn-primary" onclick="openCreateModal()">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Nuevo Álbum</span>
          </button>
        </div>

        <div class="admin-card">
          <div class="data-table-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 50px;">ID</th>
                  <th style="width: 70px;">Carátula</th>
                  <th>Título del Álbum</th>
                  <th>Año de Publicación</th>
                  <th>Total Pistas</th>
                  <th style="width: 110px; text-align: center;">Acciones</th>
                </tr>
              </thead>
              <tbody>
                <%
                If Not rsAlbums.EOF Then
                    Do While Not rsAlbums.EOF
                        Dim albImg
                        albImg = ResolveUrl(rsAlbums("ALBUM_IMAGE_PATH"))
                %>
                  <tr>
                    <td><strong>#<%=rsAlbums("ID_ALBUM")%></strong></td>
                    <td>
                      <img src="<%=albImg%>" alt="Cover" style="width: 44px; height: 44px; border-radius: 4px; object-fit: cover;" onerror="this.src='img/default-cover.svg'">
                    </td>
                    <td style="color:#fff; font-weight: 600;"><%=SafeHtml(rsAlbums("TITLE_ALBUM"))%></td>
                    <td><%=SafeHtml(rsAlbums("RELEASE_YEAR"))%></td>
                    <td><%=rsAlbums("TRACK_COUNT")%> canciones</td>
                    <td>
                      <div class="actions-cell" style="justify-content: center;">
                        <button type="button" class="btn-icon" title="Editar Álbum"
                                data-id="<%=rsAlbums("ID_ALBUM")%>"
                                data-title="<%=SafeHtml(rsAlbums("TITLE_ALBUM"))%>"
                                data-year="<%=rsAlbums("RELEASE_YEAR")%>"
                                data-image="<%=SafeHtml(rsAlbums("ALBUM_IMAGE_PATH"))%>"
                                onclick="handleEditAlbum(this)">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                        </button>
                        <button type="button" class="btn-icon delete" title="Eliminar Álbum"
                                data-id="<%=rsAlbums("ID_ALBUM")%>"
                                data-title="<%=SafeHtml(rsAlbums("TITLE_ALBUM"))%>"
                                onclick="handleDeleteAlbum(this)">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                        </button>
                      </div>
                    </td>
                  </tr>
                <%
                        rsAlbums.MoveNext
                    Loop
                Else
                    Response.Write "<tr><td colspan=""6"" style=""text-align:center; padding:24px;"">No hay álbumes registrados.</td></tr>"
                End If
                CloseRecordset rsAlbums
                %>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </main>
  </div>

  <!-- Modal Crear Álbum -->
  <div id="modalCreateAlbum" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Nuevo Álbum</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalCreateAlbum')">&times;</button>
      </div>
      <form action="admin_albums.asp" method="POST">
        <input type="hidden" name="action" value="create">

        <div class="form-group">
          <label class="form-label" for="newAlbumTitle">Título del Álbum *</label>
          <input type="text" id="newAlbumTitle" name="title_album" class="form-control" placeholder="Ej. A Night at the Opera" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="newAlbumYear">Año de Lanzamiento</label>
          <input type="number" id="newAlbumYear" name="release_year" class="form-control" value="2026" min="1900" max="2100">
        </div>

        <div class="form-group">
          <label class="form-label" for="newAlbumImage">Carátula del Álbum</label>
          <select id="newAlbumImage" name="album_image_path" class="form-control">
            <option value="~/albumImages/album1.png">~/albumImages/album1.png (Rock Album)</option>
            <option value="~/albumImages/album2.png">~/albumImages/album2.png (Relax Music)</option>
            <option value="~/albumImages/album3.png">~/albumImages/album3.png (Retro Album)</option>
            <option value="~/albumImages/album4.png">~/albumImages/album4.png (Classic Album)</option>
            <option value="~/albumImages/album5.png">~/albumImages/album5.png (Country Album)</option>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalCreateAlbum')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Guardar Álbum</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Modal Editar Álbum -->
  <div id="modalEditAlbum" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Editar Álbum</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalEditAlbum')">&times;</button>
      </div>
      <form action="admin_albums.asp" method="POST">
        <input type="hidden" name="action" value="edit">
        <input type="hidden" id="editAlbumId" name="id_album" value="">

        <div class="form-group">
          <label class="form-label" for="editAlbumTitle">Título del Álbum *</label>
          <input type="text" id="editAlbumTitle" name="title_album" class="form-control" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="editAlbumYear">Año de Lanzamiento</label>
          <input type="number" id="editAlbumYear" name="release_year" class="form-control" min="1900" max="2100">
        </div>

        <div class="form-group">
          <label class="form-label" for="editAlbumImage">Carátula del Álbum</label>
          <input type="text" id="editAlbumImage" name="album_image_path" class="form-control" list="albumImagesList" maxlength="255">
          <datalist id="albumImagesList">
            <option value="~/albumImages/album1.png">~/albumImages/album1.png (Rock Album)</option>
            <option value="~/albumImages/album2.png">~/albumImages/album2.png (Relax Music)</option>
            <option value="~/albumImages/album3.png">~/albumImages/album3.png (Retro Album)</option>
            <option value="~/albumImages/album4.png">~/albumImages/album4.png (Classic Album)</option>
            <option value="~/albumImages/album5.png">~/albumImages/album5.png (Country Album)</option>
          </datalist>
          <span style="font-size:11px; color: var(--text-dim); margin-top:4px; display:block;">Puedes seleccionar una carátula existente en <code>AlbumImages/</code> o escribir la ruta de una nueva.</span>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalEditAlbum')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Actualizar Álbum</button>
        </div>
      </form>
    </div>
  </div>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
  <script>
    function openCreateModal() {
      openModal('modalCreateAlbum');
    }

    function handleEditAlbum(btn) {
      document.getElementById('editAlbumId').value = btn.getAttribute('data-id');
      document.getElementById('editAlbumTitle').value = btn.getAttribute('data-title');
      document.getElementById('editAlbumYear').value = btn.getAttribute('data-year');
      document.getElementById('editAlbumImage').value = btn.getAttribute('data-image');
      openModal('modalEditAlbum');
    }

    function handleDeleteAlbum(btn) {
      var id = btn.getAttribute('data-id');
      var title = btn.getAttribute('data-title');
      confirmAction('¿Eliminar el álbum \'' + title + '\'?', 'admin_albums.asp?action=delete&id=' + id);
    }

    function openEditModal(id, title, year, image) {
      document.getElementById('editAlbumId').value = id;
      document.getElementById('editAlbumTitle').value = title;
      document.getElementById('editAlbumYear').value = year;
      document.getElementById('editAlbumImage').value = image;
      openModal('modalEditAlbum');
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
