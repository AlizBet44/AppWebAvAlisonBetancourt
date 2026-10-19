<%@ Language="VBScript" CodePage=65001 %>
<!--#include file="includes/conexion.asp"-->
<%
' Si ya está autenticado, redirigir al index
If Not IsEmpty(Session("UserId")) And Trim(Session("UserId") & "") <> "" Then
    Response.Redirect "index.asp"
    Response.End
End If

Dim errorMsg
errorMsg = ""

' Procesar Registro de Nuevo Usuario (POST)
If Request.ServerVariables("REQUEST_METHOD") = "POST" Then
    Dim txtUsername, txtEmail, txtPassword, txtConfirmPass, txtBirthday
    txtUsername    = Trim(Request.Form("username"))
    txtEmail       = Trim(Request.Form("email"))
    txtPassword    = Trim(Request.Form("password"))
    txtConfirmPass = Trim(Request.Form("confirm_password"))
    txtBirthday    = Trim(Request.Form("birthday"))

    ' Validaciones de Servidor
    If txtUsername = "" Or txtEmail = "" Or txtPassword = "" Then
        errorMsg = "Por favor, completa todos los campos requeridos."
    ElseIf txtPassword <> txtConfirmPass Then
        errorMsg = "Las contraseñas no coinciden. Por favor verifícalas."
    ElseIf Len(txtPassword) < 6 Then
        errorMsg = "La contraseña debe tener al menos 6 caracteres."
    Else
        Dim conn, cmdCheck, rsCheck
        On Error Resume Next
        Set conn = GetConnection()

        If Err.Number <> 0 Then
            errorMsg = "Error al conectar con la base de datos: " & Err.Description
            Err.Clear
        Else
            ' 1. Verificar si el usuario o email ya existen (Consulta Parametrizada)
            Set cmdCheck = Server.CreateObject("ADODB.Command")
            cmdCheck.ActiveConnection = conn
            cmdCheck.CommandType = adCmdText
            cmdCheck.CommandText = "SELECT COUNT(*) AS Total FROM TBL_USER WHERE USERNAME = ? OR EMAIL = ?"
            cmdCheck.Parameters.Append cmdCheck.CreateParameter("@USERNAME", adVarChar, adParamInput, 50, txtUsername)
            cmdCheck.Parameters.Append cmdCheck.CreateParameter("@EMAIL", adVarChar, adParamInput, 50, txtEmail)
            
            Set rsCheck = cmdCheck.Execute()

            If Not rsCheck.EOF And CLng(rsCheck("Total")) > 0 Then
                errorMsg = "El nombre de usuario o el correo electrónico ya se encuentran registrados."
                CloseRecordset rsCheck
            Else
                CloseRecordset rsCheck

                ' 2. Insertar nuevo usuario con rol Normal (USER_TYPE = 0) de forma parametrizada
                Dim cmdInsert
                Set cmdInsert = Server.CreateObject("ADODB.Command")
                cmdInsert.ActiveConnection = conn
                cmdInsert.CommandType = adCmdText
                cmdInsert.CommandText = "INSERT INTO TBL_USER (USERNAME, PASSWORD, EMAIL, BIRTHDAY, USER_PHOTO, USER_TYPE) VALUES (?, ?, ?, ?, ?, 0)"

                cmdInsert.Parameters.Append cmdInsert.CreateParameter("@USERNAME", adVarChar, adParamInput, 50, txtUsername)
                cmdInsert.Parameters.Append cmdInsert.CreateParameter("@PASSWORD", adVarChar, adParamInput, 50, txtPassword)
                cmdInsert.Parameters.Append cmdInsert.CreateParameter("@EMAIL", adVarChar, adParamInput, 50, txtEmail)

                ' Manejo de fecha de nacimiento opcional
                If IsDate(txtBirthday) Then
                    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@BIRTHDAY", adDate, adParamInput, , CDate(txtBirthday))
                Else
                    cmdInsert.Parameters.Append cmdInsert.CreateParameter("@BIRTHDAY", adDate, adParamInput, , Null)
                End If

                cmdInsert.Parameters.Append cmdInsert.CreateParameter("@USER_PHOTO", adVarChar, adParamInput, 255, "~/userImages/default.jpg")

                cmdInsert.Execute()

                If Err.Number <> 0 Then
                    errorMsg = "Error al registrar la cuenta: " & Err.Description
                    Err.Clear
                Else
                    CloseConnection conn
                    Response.Redirect "login.asp?msg=registered"
                    Response.End
                End If
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
  <title>Crear Cuenta • PoliMusic</title>
  <link rel="stylesheet" href="css/app.css">
  <link rel="stylesheet" href="css/auth.css">
</head>
<body class="auth-body">
  <div class="auth-ambient-glow"></div>

  <div class="auth-container">
    <div class="auth-card">
      <div class="auth-header">
        <a href="login.asp">
          <img src="img/logo.svg" alt="PoliMusic Logo" class="auth-logo" style="height: 44px;">
        </a>
        <h1 class="auth-title">Regístrate en PoliMusic</h1>
        <p class="auth-subtitle">Crea tu cuenta institucional y empieza a escuchar</p>
      </div>

      <% If errorMsg <> "" Then %>
        <div class="alert alert-danger">
          <svg viewBox="0 0 24 24" width="18" height="18" fill="currentColor"><path d="M12 2C6.48 2 2 6.48 2 12s4.48 10 10 10 10-4.48 10-10S17.52 2 12 2zm1 15h-2v-2h2v2zm0-4h-2V7h2v6z"/></svg>
          <span><%=SafeHtml(errorMsg)%></span>
        </div>
      <% End If %>

      <form action="register.asp" method="POST" class="auth-form" autocomplete="off">
        <div class="form-group">
          <label class="form-label" for="regUsername">Nombre de Usuario *</label>
          <div class="input-icon-group">
            <input type="text" id="regUsername" name="username" class="form-control" placeholder="Ej. usuario.epn" required autofocus maxlength="50" value="<%=SafeHtml(Request.Form("username"))%>">
            <svg viewBox="0 0 24 24"><path d="M12 12c2.21 0 4-1.79 4-4s-1.79-4-4-4-4 1.79-4 4 1.79 4 4 4zm0 2c-2.67 0-8 1.34-8 4v2h16v-2c0-2.66-5.33-4-8-4z"/></svg>
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="regEmail">Correo Electrónico *</label>
          <div class="input-icon-group">
            <input type="email" id="regEmail" name="email" class="form-control" placeholder="tu.correo@epn.edu.ec" required maxlength="50" value="<%=SafeHtml(Request.Form("email"))%>">
            <svg viewBox="0 0 24 24"><path d="M20 4H4c-1.1 0-1.99.9-1.99 2L2 18c0 1.1.9 2 2 2h16c1.1 0 2-.9 2-2V6c0-1.1-.9-2-2-2zm0 4l-8 5-8-5V6l8 5 8-5v2z"/></svg>
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="regPassword">Contraseña *</label>
          <div class="input-icon-group">
            <input type="password" id="regPassword" name="password" class="form-control" placeholder="Mínimo 6 caracteres" required minlength="6">
            <svg viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="regConfirmPass">Confirmar Contraseña *</label>
          <div class="input-icon-group">
            <input type="password" id="regConfirmPass" name="confirm_password" class="form-control" placeholder="Repite tu contraseña" required minlength="6">
            <svg viewBox="0 0 24 24"><path d="M18 8h-1V6c0-2.76-2.24-5-5-5S7 3.24 7 6v2H6c-1.1 0-2 .9-2 2v10c0 1.1.9 2 2 2h12c1.1 0 2-.9 2-2V10c0-1.1-.9-2-2-2zm-6 9c-1.1 0-2-.9-2-2s.9-2 2-2 2 .9 2 2-.9 2-2 2zm3.1-9H8.9V6c0-1.71 1.39-3.1 3.1-3.1 1.71 0 3.1 1.39 3.1 3.1v2z"/></svg>
          </div>
        </div>

        <div class="form-group">
          <label class="form-label" for="regBirthday">Fecha de Nacimiento (Opcional)</label>
          <div class="input-icon-group">
            <input type="date" id="regBirthday" name="birthday" class="form-control" value="<%=SafeHtml(Request.Form("birthday"))%>">
            <svg viewBox="0 0 24 24"><path d="M19 3h-1V1h-2v2H8V1H6v2H5c-1.11 0-1.99.9-1.99 2L3 19c0 1.1.89 2 2 2h14c1.1 0 2-.9 2-2V5c0-1.1-.9-2-2-2zm0 16H5V8h14v11zM7 10h5v5H7z"/></svg>
          </div>
        </div>

        <button type="submit" class="btn btn-primary auth-btn-submit">
          Crear mi Cuenta
        </button>
      </form>

      <div class="auth-footer">
        ¿Ya tienes una cuenta registrada?
        <a href="login.asp" class="auth-link">Inicia sesión</a>
      </div>
    </div>
  </div>
</body>
</html>
