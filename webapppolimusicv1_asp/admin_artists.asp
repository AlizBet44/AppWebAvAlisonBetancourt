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

' OPERACIÓN: CREAR ARTISTA
If reqMethod = "POST" And action = "create" Then
    Dim newName, newLast, newImg
    newName = Trim(Request.Form("artist_name"))
    newLast = Trim(Request.Form("artist_lastname"))
    newImg  = Trim(Request.Form("artist_image"))

    If newName = "" Then
        errorMsg = "El nombre del artista es obligatorio."
    Else
        If newImg = "" Then newImg = "~/ArtistImages/artist1.jpg"

        On Error Resume Next
        Dim cmdInsArt
        Set cmdInsArt = Server.CreateObject("ADODB.Command")
        cmdInsArt.ActiveConnection = conn
        cmdInsArt.CommandType = adCmdText
        cmdInsArt.CommandText = "INSERT INTO TBL_ARTIST (ARTIST_NAME, ARTIST_LASTNAME, ARTIST_IMAGE) VALUES (?, ?, ?)"
        cmdInsArt.Parameters.Append cmdInsArt.CreateParameter("@NAME", adVarChar, adParamInput, 100, newName)
        cmdInsArt.Parameters.Append cmdInsArt.CreateParameter("@LAST", adVarChar, adParamInput, 100, newLast)
        cmdInsArt.Parameters.Append cmdInsArt.CreateParameter("@IMG", adVarChar, adParamInput, 255, newImg)
        cmdInsArt.Execute()

        If Err.Number <> 0 Then
            errorMsg = "Error al crear el artista: " & Err.Description
            Err.Clear
        Else
            infoMsg = "Artista '" & newName & " " & newLast & "' registrado exitosamente."
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: EDITAR ARTISTA
ElseIf reqMethod = "POST" And action = "edit" Then
    Dim editId, editName, editLast, editImg
    editId   = Request.Form("id_artist")
    editName = Trim(Request.Form("artist_name"))
    editLast = Trim(Request.Form("artist_lastname"))
    editImg  = Trim(Request.Form("artist_image"))

    If Not IsNumeric(editId) Or editName = "" Then
        errorMsg = "Datos incompletos para actualizar el artista."
    Else
        If editImg = "" Then editImg = "~/ArtistImages/artist1.jpg"
        On Error Resume Next
        Dim cmdUpdArt
        Set cmdUpdArt = Server.CreateObject("ADODB.Command")
        cmdUpdArt.ActiveConnection = conn
        cmdUpdArt.CommandType = adCmdText
        cmdUpdArt.CommandText = "UPDATE TBL_ARTIST SET ARTIST_NAME = ?, ARTIST_LASTNAME = ?, ARTIST_IMAGE = ? WHERE ID_ARTIST = ?"
        cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@NAME", adVarChar, adParamInput, 100, editName)
        cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@LAST", adVarChar, adParamInput, 100, editLast)
        cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@IMG", adVarChar, adParamInput, 255, editImg)
        cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@ID", adInteger, adParamInput, , CLng(editId))
        cmdUpdArt.Execute()

        If Err.Number <> 0 Then
            errorMsg = "Error al actualizar el artista: " & Err.Description
            Err.Clear
        Else
            infoMsg = "Artista actualizado correctamente."
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: ELIMINAR ARTISTA
ElseIf action = "delete" Then
    Dim delId
    delId = Request("id")
    If IsNumeric(delId) Then
        On Error Resume Next
        Dim cmdDelArt
        Set cmdDelArt = Server.CreateObject("ADODB.Command")
        cmdDelArt.ActiveConnection = conn
        cmdDelArt.CommandType = adCmdText
        cmdDelArt.CommandText = "DELETE FROM TBL_PLAYER WHERE ID_ARTIST = ?; DELETE FROM TBL_ARTIST WHERE ID_ARTIST = ?;"
        cmdDelArt.Parameters.Append cmdDelArt.CreateParameter("@P1", adInteger, adParamInput, , CLng(delId))
        cmdDelArt.Parameters.Append cmdDelArt.CreateParameter("@P2", adInteger, adParamInput, , CLng(delId))
        cmdDelArt.Execute()

        infoMsg = "Artista eliminado correctamente."
        On Error GoTo 0
    End If
End If

' Listado de artistas
Dim rsArtists
Set rsArtists = conn.Execute("SELECT a.ID_ARTIST, a.ARTIST_NAME, a.ARTIST_LASTNAME, a.ARTIST_IMAGE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYER WHERE ID_ARTIST = a.ID_ARTIST) AS SONG_COUNT " & _
    "FROM TBL_ARTIST a ORDER BY a.ID_ARTIST DESC")
%>
    <!DOCTYPE html>
    <html lang="es">

    <head>
      <meta charset="UTF-8">
      <meta name="viewport" content="width=device-width, initial-scale=1.0">
      <title>CRUD Artistas • PoliMusic Admin</title>
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
                <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                  <path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z" />
                </svg>
                <span>
                  <%=SafeHtml(infoMsg)%>
                </span>
              </div>
              <% End If %>

                <% If errorMsg <> "" Then %>
                  <div class="alert alert-danger">
                    <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor">
                      <path
                        d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z" />
                    </svg>
                    <span>
                      <%=SafeHtml(errorMsg)%>
                    </span>
                  </div>
                  <% End If %>

                    <div class="section-header">
                      <div>
                        <h1 class="section-title">Gestión de Artistas (CRUD)</h1>
                        <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Crear, editar
                          información de intérpretes y gestionar perfiles artísticos</p>
                      </div>
                      <button type="button" class="btn btn-primary" onclick="openCreateModal()">
                        <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor">
                          <path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z" />
                        </svg>
                        <span>Nuevo Artista</span>
                      </button>
                    </div>

                    <div class="admin-card">
                      <div class="data-table-wrapper">
                        <table class="data-table">
                          <thead>
                            <tr>
                              <th style="width: 50px;">ID</th>
                              <th style="width: 70px;">Foto</th>
                              <th>Nombre</th>
                              <th>Apellido</th>
                              <th>Canciones</th>
                              <th style="width: 110px; text-align: center;">Acciones</th>
                            </tr>
                          </thead>
                          <tbody>
                            <%
                            If Not rsArtists.EOF Then
                                Do While Not rsArtists.EOF
                                    Dim artImg
                                    artImg = ResolveUrl(rsArtists("ARTIST_IMAGE"))
                            %>
                              <tr>
                                <td><strong>#<%=rsArtists("ID_ARTIST")%></strong></td>
                                <td>
                                  <img src="<%=artImg%>" alt="Foto"
                                    style="width: 44px; height: 44px; border-radius: 50%; object-fit: cover; border: 2px solid var(--primary);"
                                    onerror="this.src='img/default-cover.svg'">
                                </td>
                                <td style="color:#fff; font-weight: 600;">
                                  <%=SafeHtml(rsArtists("ARTIST_NAME"))%>
                                </td>
                                <td>
                                  <%=SafeHtml(rsArtists("ARTIST_LASTNAME"))%>
                                </td>
                                <td>
                                  <%=rsArtists("SONG_COUNT")%> canciones
                                </td>
                                <td>
                                  <div class="actions-cell" style="justify-content: center;">
                                    <button type="button" class="btn-icon" title="Editar Artista"
                                      data-id="<%=rsArtists("ID_ARTIST")%>"
                                      data-name="<%=SafeHtml(rsArtists("ARTIST_NAME"))%>"
                                      data-lastname="<%=SafeHtml(rsArtists("ARTIST_LASTNAME"))%>"
                                      data-image="<%=SafeHtml(rsArtists("ARTIST_IMAGE"))%>"
                                      onclick="handleEditArtist(this)">
                                      <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor">
                                        <path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z" />
                                      </svg>
                                    </button>
                                    <button type="button" class="btn-icon delete" title="Eliminar Artista"
                                      data-id="<%=rsArtists("ID_ARTIST")%>"
                                      data-name="<%=SafeHtml(rsArtists("ARTIST_NAME"))%>"
                                      onclick="handleDeleteArtist(this)">
                                      <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor">
                                        <path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z" />
                                      </svg>
                                    </button>
                                  </div>
                                </td>
                              </tr>
                            <%
                                    rsArtists.MoveNext
                                Loop
                            Else
                                Response.Write "<tr><td colspan='6' style='text-align:center; padding:24px;'>No hay artistas registrados.</td></tr>"
                            End If
                            CloseRecordset rsArtists
                            %>
                          </tbody>
                        </table>
                      </div>
                    </div>
          </div>
        </main>
      </div>

      <!-- Modal Crear Artista -->
      <div id="modalCreateArtist" class="modal-overlay">
        <div class="modal-content">
          <div class="modal-header">
            <h3 class="modal-title">Nuevo Artista</h3>
            <button type="button" class="modal-close" onclick="closeModal('modalCreateArtist')">&times;</button>
          </div>
          <form action="admin_artists.asp" method="POST">
            <input type="hidden" name="action" value="create">

            <div class="form-group">
              <label class="form-label" for="newArtistName">Nombre *</label>
              <input type="text" id="newArtistName" name="artist_name" class="form-control" placeholder="Ej. Juan"
                required maxlength="100">
            </div>

            <div class="form-group">
              <label class="form-label" for="newArtistLast">Apellido / Sobrenombre</label>
              <input type="text" id="newArtistLast" name="artist_lastname" class="form-control"
                placeholder="Ej. Gabriel" maxlength="100">
            </div>

            <div class="form-group">
              <label class="form-label" for="newArtistImage">Foto del Artista</label>
              <select id="newArtistImage" name="artist_image" class="form-control">
                <option value="~/ArtistImages/artist1.jpg">~/ArtistImages/artist1.jpg (Juan Gabriel)</option>
                <option value="~/ArtistImages/artist2.jpg">~/ArtistImages/artist2.jpg (Ana Gabriel)</option>
                <option value="~/ArtistImages/artist3.jpg">~/ArtistImages/artist3.jpg (Mariah Carey)</option>
                <option value="~/ArtistImages/artist4.jpg">~/ArtistImages/artist4.jpg (Elvis Presley)</option>
                <option value="~/ArtistImages/artist5.png">~/ArtistImages/artist5.png (Ed Sheeran)</option>
              </select>
            </div>

            <div class="modal-actions">
              <button type="button" class="btn btn-secondary"
                onclick="closeModal('modalCreateArtist')">Cancelar</button>
              <button type="submit" class="btn btn-primary">Guardar Artista</button>
            </div>
          </form>
        </div>
      </div>

      <!-- Modal Editar Artista -->
      <div id="modalEditArtist" class="modal-overlay">
        <div class="modal-content">
          <div class="modal-header">
            <h3 class="modal-title">Editar Artista</h3>
            <button type="button" class="modal-close" onclick="closeModal('modalEditArtist')">&times;</button>
          </div>
          <form action="admin_artists.asp" method="POST">
            <input type="hidden" name="action" value="edit">
            <input type="hidden" id="editArtistId" name="id_artist" value="">

            <div class="form-group">
              <label class="form-label" for="editArtistName">Nombre *</label>
              <input type="text" id="editArtistName" name="artist_name" class="form-control" required maxlength="100">
            </div>

            <div class="form-group">
              <label class="form-label" for="editArtistLast">Apellido / Sobrenombre</label>
              <input type="text" id="editArtistLast" name="artist_lastname" class="form-control" maxlength="100">
            </div>

            <div class="form-group">
              <label class="form-label" for="editArtistImage">Foto del Artista</label>
              <input type="text" id="editArtistImage" name="artist_image" class="form-control" list="artistImagesList" maxlength="255">
              <datalist id="artistImagesList">
                <option value="~/ArtistImages/artist1.jpg">~/ArtistImages/artist1.jpg (Juan Gabriel)</option>
                <option value="~/ArtistImages/artist2.jpg">~/ArtistImages/artist2.jpg (Ana Gabriel)</option>
                <option value="~/ArtistImages/artist3.jpg">~/ArtistImages/artist3.jpg (Mariah Carey)</option>
                <option value="~/ArtistImages/artist4.jpg">~/ArtistImages/artist4.jpg (Elvis Presley)</option>
                <option value="~/ArtistImages/artist5.png">~/ArtistImages/artist5.png (Ed Sheeran)</option>
              </datalist>
              <span style="font-size:11px; color: var(--text-dim); margin-top:4px; display:block;">Puedes seleccionar una foto existente en <code>ArtistImages/</code> o ingresar la ruta de una foto nueva.</span>
            </div>

            <div class="modal-actions">
              <button type="button" class="btn btn-secondary" onclick="closeModal('modalEditArtist')">Cancelar</button>
              <button type="submit" class="btn btn-primary">Actualizar Artista</button>
            </div>
          </form>
        </div>
      </div>

      <!--#include file="includes/player.asp"-->

      <script src="js/app.js"></script>
      <script src="js/player.js"></script>
      <script>
        function openCreateModal() {
          openModal('modalCreateArtist');
        }

        function openEditModal(id, name, lastname, img) {
          document.getElementById('editArtistId').value = id;
          document.getElementById('editArtistName').value = name;
          document.getElementById('editArtistLast').value = lastname;
          document.getElementById('editArtistImage').value = img;
          openModal('modalEditArtist');
        }

        function handleEditArtist(btn) {
          openEditModal(
            btn.getAttribute('data-id'),
            btn.getAttribute('data-name'),
            btn.getAttribute('data-lastname'),
            btn.getAttribute('data-image')
          );
        }

        function handleDeleteArtist(btn) {
          var id = btn.getAttribute('data-id');
          var name = btn.getAttribute('data-name');
          confirmAction('¿Eliminar al artista \'' + name + '\'?', 'admin_artists.asp?action=delete&id=' + encodeURIComponent(id));
        }
      </script>
    </body>

    </html>
    <% CloseConnection conn %>