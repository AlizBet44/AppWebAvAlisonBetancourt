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

' OPERACIÓN: CREAR USUARIO
If reqMethod = "POST" And action = "create" Then
    Dim newUsername, newPassword, newEmail, newBirthday, newPhoto, newType
    newUsername = Trim(Request.Form("username"))
    newPassword = Trim(Request.Form("password"))
    newEmail    = Trim(Request.Form("email"))
    newBirthday = Trim(Request.Form("birthday"))
    newPhoto    = Trim(Request.Form("user_photo"))
    newType     = Request.Form("user_type")

    If newUsername = "" Or newPassword = "" Or newEmail = "" Then
        errorMsg = "Nombre de usuario, correo y contraseña son obligatorios."
    Else
        If Not IsNumeric(newType) Then newType = 0
        If newPhoto = "" Then newPhoto = "~/userImages/default.jpg"

        On Error Resume Next
        ' Verificar duplicados
        Dim cmdChk, rsChk
        Set cmdChk = Server.CreateObject("ADODB.Command")
        cmdChk.ActiveConnection = conn
        cmdChk.CommandType = adCmdText
        cmdChk.CommandText = "SELECT COUNT(*) AS Total FROM TBL_USER WHERE USERNAME = ? OR EMAIL = ?"
        cmdChk.Parameters.Append cmdChk.CreateParameter("@U", adVarChar, adParamInput, 50, newUsername)
        cmdChk.Parameters.Append cmdChk.CreateParameter("@E", adVarChar, adParamInput, 50, newEmail)
        Set rsChk = cmdChk.Execute()

        If Not rsChk.EOF And CLng(rsChk("Total")) > 0 Then
            errorMsg = "El nombre de usuario o correo ya existen."
            CloseRecordset rsChk
        Else
            CloseRecordset rsChk

            ' Insertar Usuario de forma parametrizada
            Dim cmdInsUser
            Set cmdInsUser = Server.CreateObject("ADODB.Command")
            cmdInsUser.ActiveConnection = conn
            cmdInsUser.CommandType = adCmdText
            cmdInsUser.CommandText = "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES (?, ?, ?, ?, ?, ?)"
            cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@U", adVarChar, adParamInput, 50, newUsername)
            cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@P", adVarChar, adParamInput, 50, newPassword)
            cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@E", adVarChar, adParamInput, 50, newEmail)

            If IsDate(newBirthday) Then
                cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@B", adDate, adParamInput, , CDate(newBirthday))
            Else
                cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@B", adDate, adParamInput, , Null)
            End If

            cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@PH", adVarChar, adParamInput, 255, newPhoto)
            cmdInsUser.Parameters.Append cmdInsUser.CreateParameter("@T", adInteger, adParamInput, , CInt(newType))
            cmdInsUser.Execute()

            If Err.Number <> 0 Then
                errorMsg = "Error al crear el usuario: " & Err.Description
                Err.Clear
            Else
                infoMsg = "Usuario '" & newUsername & "' creado exitosamente."
            End If
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: EDITAR USUARIO
ElseIf reqMethod = "POST" And action = "edit" Then
    Dim editId, editUsername, editPassword, editEmail, editBirthday, editPhoto, editType
    editId       = Request.Form("id_user")
    editUsername = Trim(Request.Form("username"))
    editPassword = Trim(Request.Form("password"))
    editEmail    = Trim(Request.Form("email"))
    editBirthday = Trim(Request.Form("birthday"))
    editPhoto    = Trim(Request.Form("user_photo"))
    editType     = Request.Form("user_type")

    If Not IsNumeric(editId) Or editUsername = "" Or editEmail = "" Then
        errorMsg = "Datos incompletos para actualizar el usuario."
    Else
        If Not IsNumeric(editType) Then editType = 0
        If editPhoto = "" Then editPhoto = "~/userImages/default.jpg"

        On Error Resume Next
        Dim cmdUpdUser
        Set cmdUpdUser = Server.CreateObject("ADODB.Command")
        cmdUpdUser.ActiveConnection = conn
        cmdUpdUser.CommandType = adCmdText

        If editPassword <> "" Then
            ' Actualiza con nueva contraseña
            cmdUpdUser.CommandText = "UPDATE TBL_USER SET USERNAME = ?, PASSWORD = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?"
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@U", adVarChar, adParamInput, 50, editUsername)
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@P", adVarChar, adParamInput, 50, editPassword)
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@E", adVarChar, adParamInput, 50, editEmail)
            If IsDate(editBirthday) Then
                cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@B", adDate, adParamInput, , CDate(editBirthday))
            Else
                cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@B", adDate, adParamInput, , Null)
            End If
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@PH", adVarChar, adParamInput, 255, editPhoto)
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@T", adInteger, adParamInput, , CInt(editType))
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@ID", adInteger, adParamInput, , CLng(editId))
        Else
            ' Mantiene la contraseña existente
            cmdUpdUser.CommandText = "UPDATE TBL_USER SET USERNAME = ?, EMAIL = ?, BIRTHDAY = ?, USER_PHOTO = ?, USER_TYPE = ? WHERE ID_USER = ?"
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@U", adVarChar, adParamInput, 50, editUsername)
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@E", adVarChar, adParamInput, 50, editEmail)
            If IsDate(editBirthday) Then
                cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@B", adDate, adParamInput, , CDate(editBirthday))
            Else
                cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@B", adDate, adParamInput, , Null)
            End If
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@PH", adVarChar, adParamInput, 255, editPhoto)
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@T", adInteger, adParamInput, , CInt(editType))
            cmdUpdUser.Parameters.Append cmdUpdUser.CreateParameter("@ID", adInteger, adParamInput, , CLng(editId))
        End If

        cmdUpdUser.Execute()

        If Err.Number <> 0 Then
            errorMsg = "Error al actualizar el usuario: " & Err.Description
            Err.Clear
        Else
            infoMsg = "Usuario actualizado correctamente."
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: ELIMINAR USUARIO
ElseIf action = "delete" Then
    Dim delId
    delId = Request("id")
    If IsNumeric(delId) Then
        If CLng(delId) = CurrentUserId() Then
            errorMsg = "Por seguridad, no puedes eliminar tu propia cuenta de administrador en sesión."
        Else
            On Error Resume Next
            ' 1. Eliminar canciones de playlists del usuario
            conn.Execute "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_PLAYLIST IN (SELECT ID_PLAYLIST FROM TBL_PLAYLIST WHERE ID_USER = " & CLng(delId) & ")"
            ' 2. Eliminar playlists del usuario
            conn.Execute "DELETE FROM TBL_PLAYLIST WHERE ID_USER = " & CLng(delId)
            ' 3. Eliminar usuario
            Dim cmdDelU
            Set cmdDelU = Server.CreateObject("ADODB.Command")
            cmdDelU.ActiveConnection = conn
            cmdDelU.CommandType = adCmdText
            cmdDelU.CommandText = "DELETE FROM TBL_USER WHERE ID_USER = ?"
            cmdDelU.Parameters.Append cmdDelU.CreateParameter("@ID", adInteger, adParamInput, , CLng(delId))
            cmdDelU.Execute()

            infoMsg = "Usuario eliminado del sistema."
            On Error GoTo 0
        End If
    End If
End If

' Listado de usuarios
Dim rsUsers
Set rsUsers = conn.Execute("SELECT ID_USER, USERNAME, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE, " & _
    "(SELECT COUNT(*) FROM TBL_PLAYLIST WHERE ID_USER = u.ID_USER) AS PLAYLIST_COUNT " & _
    "FROM TBL_USER u ORDER BY u.ID_USER ASC")
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CRUD Usuarios • PoliMusic Admin</title>
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
            <h1 class="section-title">Gestión de Usuarios y Roles (RBAC)</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Crear usuarios, modificar contraseñas y asignar roles de Administrador o Normal</p>
          </div>
          <button type="button" class="btn btn-primary" onclick="openCreateModal()">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Nuevo Usuario</span>
          </button>
        </div>

        <div class="admin-card">
          <div class="data-table-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 50px;">ID</th>
                  <th style="width: 60px;">Foto</th>
                  <th>Usuario</th>
                  <th>Correo Electrónico</th>
                  <th>Fecha Nacimiento</th>
                  <th>Rol Asignado</th>
                  <th>Playlists</th>
                  <th style="width: 110px; text-align: center;">Acciones</th>
                </tr>
              </thead>
              <tbody>
                <%
                If Not rsUsers.EOF Then
                    Do While Not rsUsers.EOF
                        Dim uPhoto, roleBadgeClass, bDayStr
                        uPhoto = ResolveUrl(rsUsers("USER_PHOTO"))
                        If CInt(rsUsers("USER_TYPE")) = 1 Then
                            roleBadgeClass = "role-tag-admin"
                        Else
                            roleBadgeClass = "role-tag-user"
                        End If
                        bDayStr = FormatDateStr(rsUsers("BIRTHDAY"))
                %>
                  <tr>
                    <td><strong>#<%=rsUsers("ID_USER")%></strong></td>
                    <td>
                      <img src="<%=uPhoto%>" alt="Avatar" style="width: 40px; height: 40px; border-radius: 50%; object-fit: cover;" onerror="this.src='UserImages/default.jpg'">
                    </td>
                    <td style="color:#fff; font-weight: 700;">
                      <%=SafeHtml(rsUsers("USERNAME"))%>
                      <% If CLng(rsUsers("ID_USER")) = CurrentUserId() Then %>
                        <span style="font-size: 11px; color: var(--accent-pink); font-weight: normal;">(Tú)</span>
                      <% End If %>
                    </td>
                    <td><%=SafeHtml(rsUsers("EMAIL"))%></td>
                    <td><%=bDayStr%></td>
                    <td>
                      <span class="role-tag <%=roleBadgeClass%>"><%=FormatRoleName(rsUsers("USER_TYPE"))%></span>
                    </td>
                    <td><%=rsUsers("PLAYLIST_COUNT")%> listas</td>
                    <td>
                      <div class="actions-cell" style="justify-content: center;">
                        <button type="button" class="btn-icon" title="Editar Usuario"
                                data-id="<%=rsUsers("ID_USER")%>"
                                data-username="<%=SafeHtml(rsUsers("USERNAME"))%>"
                                data-email="<%=SafeHtml(rsUsers("EMAIL"))%>"
                                data-type="<%=rsUsers("USER_TYPE")%>"
                                data-photo="<%=SafeHtml(rsUsers("USER_PHOTO"))%>"
                                onclick="handleEditUser(this)">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                        </button>
                        <% If CLng(rsUsers("ID_USER")) <> CurrentUserId() Then %>
                          <button type="button" class="btn-icon delete" title="Eliminar Usuario"
                                  data-id="<%=rsUsers("ID_USER")%>"
                                  data-username="<%=SafeHtml(rsUsers("USERNAME"))%>"
                                  onclick="handleDeleteUser(this)">
                            <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                          </button>
                        <% Else %>
                          <span style="display:inline-block; width:32px;"></span>
                        <% End If %>
                      </div>
                    </td>
                  </tr>
                <%
                        rsUsers.MoveNext
                    Loop
                End If
                CloseRecordset rsUsers
                %>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </main>
  </div>

  <!-- Modal Crear Usuario -->
  <div id="modalCreateUser" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Nuevo Usuario</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalCreateUser')">&times;</button>
      </div>
      <form action="admin_users.asp" method="POST">
        <input type="hidden" name="action" value="create">

        <div class="form-group">
          <label class="form-label" for="newUName">Nombre de Usuario *</label>
          <input type="text" id="newUName" name="username" class="form-control" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="newUEmail">Correo Electrónico *</label>
          <input type="email" id="newUEmail" name="email" class="form-control" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="newUPass">Contraseña *</label>
          <input type="password" id="newUPass" name="password" class="form-control" required minlength="6">
        </div>

        <div class="form-group">
          <label class="form-label" for="newUType">Rol en el Sistema (RBAC)</label>
          <select id="newUType" name="user_type" class="form-control">
            <option value="0">Usuario Normal (epn.fis) - Solo Lectura y Playlists Propias</option>
            <option value="1">Administrador (userAdmin) - Control CRUD Total</option>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="newUBirth">Fecha de Nacimiento</label>
          <input type="date" id="newUBirth" name="birthday" class="form-control">
        </div>

        <div class="form-group">
          <label class="form-label" for="newUPhoto">Foto de Perfil</label>
          <select id="newUPhoto" name="user_photo" class="form-control">
            <option value="~/userImages/default.jpg">~/userImages/default.jpg</option>
            <option value="~/userImages/root.png">~/userImages/root.png (Admin Avatar)</option>
            <option value="~/userImages/username1.png">~/userImages/username1.png</option>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalCreateUser')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Crear Usuario</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Modal Editar Usuario -->
  <div id="modalEditUser" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Editar Usuario</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalEditUser')">&times;</button>
      </div>
      <form action="admin_users.asp" method="POST">
        <input type="hidden" name="action" value="edit">
        <input type="hidden" id="editUserId" name="id_user" value="">

        <div class="form-group">
          <label class="form-label" for="editUName">Nombre de Usuario *</label>
          <input type="text" id="editUName" name="username" class="form-control" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="editUEmail">Correo Electrónico *</label>
          <input type="email" id="editUEmail" name="email" class="form-control" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="editUPass">Contraseña (Dejar vacío para no modificarla)</label>
          <input type="password" id="editUPass" name="password" class="form-control" placeholder="••••••••">
        </div>

        <div class="form-group">
          <label class="form-label" for="editUType">Rol en el Sistema (RBAC)</label>
          <select id="editUType" name="user_type" class="form-control">
            <option value="0">Usuario Normal (epn.fis) - Solo Lectura y Playlists Propias</option>
            <option value="1">Administrador (userAdmin) - Control CRUD Total</option>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="editUPhoto">Foto de Perfil</label>
          <input type="text" id="editUPhoto" name="user_photo" class="form-control" maxlength="255">
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalEditUser')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Actualizar Usuario</button>
        </div>
      </form>
    </div>
  </div>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
  <script>
    function openCreateModal() {
      openModal('modalCreateUser');
    }

    function handleEditUser(btn) {
      document.getElementById('editUserId').value = btn.getAttribute('data-id');
      document.getElementById('editUName').value = btn.getAttribute('data-username');
      document.getElementById('editUEmail').value = btn.getAttribute('data-email');
      document.getElementById('editUType').value = btn.getAttribute('data-type');
      document.getElementById('editUPhoto').value = btn.getAttribute('data-photo');
      document.getElementById('editUPass').value = '';
      openModal('modalEditUser');
    }

    function handleDeleteUser(btn) {
      var id = btn.getAttribute('data-id');
      var username = btn.getAttribute('data-username');
      confirmAction('¿Eliminar al usuario \'' + username + '\'?', 'admin_users.asp?action=delete&id=' + id);
    }

    function openEditModal(id, username, email, type, photo) {
      document.getElementById('editUserId').value = id;
      document.getElementById('editUName').value = username;
      document.getElementById('editUEmail').value = email;
      document.getElementById('editUType').value = type;
      document.getElementById('editUPhoto').value = photo;
      document.getElementById('editUPass').value = '';
      openModal('modalEditUser');
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
