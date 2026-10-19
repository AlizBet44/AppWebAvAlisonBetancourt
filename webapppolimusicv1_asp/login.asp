<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<%
' Si ya está autenticado, redirigir al index
If Not IsEmpty(Session("UserId")) And Trim(Session("UserId") & "") <> "" Then
    Response.Redirect "index.asp"
    Response.End
End If

Dim errorMsg, successMsg
errorMsg = ""
successMsg = ""

' Mensajes recibidos por QueryString
Dim msgCode
msgCode = Request.QueryString("msg")
If msgCode = "logged_out" Then
    successMsg = "Has cerrado sesión correctamente. ¡Esperamos verte pronto!"
ElseIf msgCode = "registered" Then
    successMsg = "¡Registro exitoso! Ya puedes iniciar sesión con tus credenciales."
ElseIf msgCode = "unauthorized" Then
    errorMsg = "Debes iniciar sesión para acceder al sistema."
End If

' Procesamiento del Formulario de Login (POST)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim txtUsername, txtPassword
    txtUsername = Trim(Request.Form("username"))
    txtPassword = Trim(Request.Form("password"))

    If txtUsername = "" Or txtPassword = "" Then
        errorMsg = "Por favor, ingresa tu usuario y contraseña."
    Else
        Dim conn, cmd, rs
        On Error Resume Next
        Set conn = GetConnection()

        If Err.Number <> 0 Then
            errorMsg = "Error al conectar con la base de datos: " & Err.Description
            Err.Clear
        Else
            ' Uso estricto de ADODB.Command con Parámetros para prevenir Inyección SQL
            Set cmd = Server.CreateObject("ADODB.Command")
            cmd.ActiveConnection = conn
            cmd.CommandType = adCmdText
            cmd.CommandText = "SELECT ID_USER, USERNAME, EMAIL, USER_TYPE, USER_PHOTO FROM TBL_USER WHERE USERNAME = ? AND PASSWORD = ?"

            ' Parámetros tipados
            cmd.Parameters.Append cmd.CreateParameter("@USERNAME", adVarChar, adParamInput, 50, txtUsername)
            cmd.Parameters.Append cmd.CreateParameter("@PASSWORD", adVarChar, adParamInput, 50, txtPassword)

            Set rs = cmd.Execute()

            If Err.Number <> 0 Then
                errorMsg = "Error en la consulta: " & Err.Description
                Err.Clear
            Else
                If Not rs.EOF Then
                    ' Autenticación Exitosa: Inicializar Variables de Sesión
                    Session("UserId")    = rs("ID_USER")
                    Session("Username")  = rs("USERNAME")
                    Session("Email")     = rs("EMAIL")
                    Session("UserType")  = rs("USER_TYPE")
                    Session("UserPhoto") = rs("USER_PHOTO")
                    
                    CloseRecordset rs
                    CloseConnection conn

                    Response.Redirect "index.asp"
                    Response.End
                Else
                    errorMsg = "Usuario o contraseña incorrectos. Verifica tus credenciales."
                End If
                CloseRecordset rs
            End If
            CloseConnection conn
        End If
        On Error GoTo 0
    End If
End If
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>Iniciar Sesión • PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/auth.css">
</head>
<body class="auth-body">
  <div class="auth-ambient-glow"></div>

  <div class="auth-container">
    <div class="auth-card">
      <div class="auth-header">
        <a href="login.asp">
          <img src="img/logo.svg" alt="PoliMusic Logo" class="auth-logo" style="height: 48px;">
        </a>
        <h1 class="auth-title">Bienvenido a PoliMusic</h1>
        <p class="auth-subtitle">Ingresa a tu cuenta para disfrutar de tu música favorita</p>
      </div>

      <% If errorMsg <> "" Then %>
        <div class="alert alert-danger">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
          <span><%=SafeHtml(errorMsg)%></span>
        </div>
      <% End If %>

      <% If successMsg <> "" Then %>
        <div class="alert alert-success">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M9 16.2L4.8 12l-1.4 1.4L9 19 21 7l-1.4-1.4L9 16.2z"/></svg>
          <span><%=SafeHtml(successMsg)%></span>
        </div>
      <% End If %>

      <form action="login.asp" method="POST" class="auth-form" autocomplete="off">
        <div class="form-group">
          <label class="form-label" for="txtUsername">Nombre de Usuario</label>
          <div class="input-icon-group">
            <input type="text" id="txtUsername" name="username" class="form-control" placeholder="Ej. userAdmin o userNormal" required autofocus>
            <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="txtPassword">Contraseña</label>
          <div class="input-icon-group">
            <input type="password" id="txtPassword" name="password" class="form-control" placeholder="••••••••" required>
            <svg viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>
          </div>
        </div>

        <button type="submit" class="btn btn-primary auth-btn-submit">
          Iniciar Sesión
        </button>
      </form>

      <!-- Accesos Rápidos de Prueba para el Evaluador -->
      <div class="test-creds-card">
        <div class="test-creds-title">
          <svg viewBox="0 0 24 24" width="14" height="14" fill="currentColor"><path d="M11 7h2v2h-2zm0 4h2v6h-2zm1-9C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm0 18c-4.41 0-8-3.59-8-8s3.59-8 8-8 8 3.59 8 8-3.59 8-8 8z"/></svg>
          <span>Credenciales de Prueba (Clic para autocompletar):</span>
        </div>
        <div class="test-creds-list">
          <div>✨ <strong>Alison Betancourt (Admin):</strong> <span class="cred-chip" onclick="fillCreds('Alison.Betancourt', '123456789')">Alison.Betancourt / 123456789</span></div>
          <div>👑 <strong>Admin Sistema:</strong> <span class="cred-chip" onclick="fillCreds('userAdmin', '12345678')">userAdmin / 12345678</span></div>
          <div>👤 <strong>Usuario Normal:</strong> <span class="cred-chip" onclick="fillCreds('userNormal', '12345678')">userNormal / 12345678</span></div>
        </div>
      </div>

      <div class="auth-footer">
        ¿No tienes una cuenta aún?
        <a href="register.asp" class="auth-link">Regístrate aquí</a>
      </div>
    </div>
  </div>

  <script>
    function fillCreds(u, p) {
      document.getElementById('txtUsername').value = u;
      document.getElementById('txtPassword').value = p;
    }
  </script>
</body>
</html>
