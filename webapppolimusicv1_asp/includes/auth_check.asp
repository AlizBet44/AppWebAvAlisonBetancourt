<%
' ==============================================================================
' PoliMusic - Control de Autenticación y Autorización Basado en Roles (RBAC)
' Roles:
'   1 = userAdmin (Administrador - Acceso total)
'   0 = epn.fis / userNormal (Usuario Normal - Solo lectura y sus propias playlists)
' ==============================================================================

' Comprueba si el usuario tiene una sesión activa válida
Function IsLoggedIn()
    If IsEmpty(Session("UserId")) Or Trim(Session("UserId") & "") = "" Then
        IsLoggedIn = False
    Else
        IsLoggedIn = True
    End If
End Function

' Comprueba si el usuario autenticado tiene rol de Administrador
Function IsAdmin()
    If IsLoggedIn() Then
        If CInt(Session("UserType")) = 1 Then
            IsAdmin = True
        Else
            IsAdmin = False
        End If
    Else
        IsAdmin = False
    End If
End Function

' Retorna el ID del usuario en sesión
Function CurrentUserId()
    If IsLoggedIn() Then
        CurrentUserId = CLng(Session("UserId"))
    Else
        CurrentUserId = 0
    End If
End Function

' Retorna el nombre de usuario
Function CurrentUsername()
    If IsLoggedIn() Then
        CurrentUsername = CStr(Session("Username"))
    Else
        CurrentUsername = "Invitado"
    End If
End Function

' Retorna la foto de perfil del usuario
Function CurrentUserPhoto()
    If IsLoggedIn() And Not IsEmpty(Session("UserPhoto")) And Trim(Session("UserPhoto") & "") <> "" Then
        CurrentUserPhoto = ResolveUrl(Session("UserPhoto"))
    Else
        CurrentUserPhoto = "UserImages/default.jpg"
    End If
End Function

' Retorna la etiqueta de rol del usuario
Function CurrentUserRole()
    If IsAdmin() Then
        CurrentUserRole = "Administrador"
    Else
        CurrentUserRole = "Estudiante EPN"
    End If
End Function

' Validador de Acceso Obligatorio: Redirecciona al login si no está autenticado
Sub RequireLogin()
    If Not IsLoggedIn() Then
        Response.Redirect "login.asp?msg=unauthorized"
        Response.End
    End If
End Sub

' Validador de Acceso Obligatorio para Administradores
Sub RequireAdmin()
    RequireLogin()
    If Not IsAdmin() Then
        Response.Redirect "index.asp?msg=forbidden"
        Response.End
    End If
End Sub
%>
