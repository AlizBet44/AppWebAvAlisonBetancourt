<%@ Language="VBScript" CodePage=65001 %>
<%
' ==============================================================================
' PoliMusic - Cierre de Sesión Seguro
' ==============================================================================
Session.Contents.RemoveAll()
Session.Abandon()
Response.Redirect "login.asp?msg=logged_out"
%>
