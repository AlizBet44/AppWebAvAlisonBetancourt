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

' ------------------------------------------------------------------------------
' 1. Procesamiento de Operaciones CRUD
' ------------------------------------------------------------------------------
Dim reqMethod, action
reqMethod = Request.ServerVariables("REQUEST_METHOD")
action    = Request("action")

' OPERACIÓN: CREAR CANCIÓN
If reqMethod = "POST" And action = "create" Then
    Dim newName, newPath, newArtistId, newAlbumId
    newName     = Trim(Request.Form("song_name"))
    newPath     = Trim(Request.Form("song_path"))
    newArtistId = Request.Form("artist_id")
    newAlbumId  = Request.Form("album_id")

    If newName = "" Or newPath = "" Then
        errorMsg = "El nombre y la ruta de la canción son obligatorios."
    Else
        On Error Resume Next
        ' Insertar Canción
        Dim cmdInsSong, rsNewId
        Set cmdInsSong = Server.CreateObject("ADODB.Command")
        cmdInsSong.ActiveConnection = conn
        cmdInsSong.CommandType = adCmdText
        cmdInsSong.CommandText = "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) VALUES (?, ?, 0); SELECT SCOPE_IDENTITY() AS NewId;"
        cmdInsSong.Parameters.Append cmdInsSong.CreateParameter("@SONG_NAME", adVarChar, adParamInput, 50, newName)
        cmdInsSong.Parameters.Append cmdInsSong.CreateParameter("@SONG_PATH", adVarChar, adParamInput, 255, newPath)
        Set rsNewId = cmdInsSong.Execute()

        If Err.Number <> 0 Then
            errorMsg = "Error al crear la canción: " & Err.Description
            Err.Clear
        Else
            Dim createdSongId
            createdSongId = rsNewId("NewId")
            CloseRecordset rsNewId

            ' Asociar Artista si fue seleccionado
            If IsNumeric(newArtistId) And CLng(newArtistId) > 0 Then
                Dim cmdArt
                Set cmdArt = Server.CreateObject("ADODB.Command")
                cmdArt.ActiveConnection = conn
                cmdArt.CommandType = adCmdText
                cmdArt.CommandText = "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)"
                cmdArt.Parameters.Append cmdArt.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(createdSongId))
                cmdArt.Parameters.Append cmdArt.CreateParameter("@ID_ARTIST", adInteger, adParamInput, , CLng(newArtistId))
                cmdArt.Execute()
            End If

            ' Asociar Álbum si fue seleccionado
            If IsNumeric(newAlbumId) And CLng(newAlbumId) > 0 Then
                Dim cmdAlb
                Set cmdAlb = Server.CreateObject("ADODB.Command")
                cmdAlb.ActiveConnection = conn
                cmdAlb.CommandType = adCmdText
                cmdAlb.CommandText = "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)"
                cmdAlb.Parameters.Append cmdAlb.CreateParameter("@ID_ALBUM", adInteger, adParamInput, , CLng(newAlbumId))
                cmdAlb.Parameters.Append cmdAlb.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(createdSongId))
                cmdAlb.Execute()
            End If

            infoMsg = "Canción '" & newName & "' agregada exitosamente al catálogo."
        End If
        On Error GoTo 0
    End If

' OPERACIÓN: EDITAR CANCIÓN
ElseIf reqMethod = "POST" And action = "edit" Then
    Dim editId, editName, editPath, editArtistId, editAlbumId
    editId       = Request.Form("id_song")
    editName     = Trim(Request.Form("song_name"))
    editPath     = Trim(Request.Form("song_path"))
    editArtistId = Request.Form("artist_id")
    editAlbumId  = Request.Form("album_id")

    If Not IsNumeric(editId) Or editName = "" Or editPath = "" Then
        errorMsg = "Datos incompletos para actualizar la canción."
    Else
        On Error Resume Next
        ' Actualizar tabla TBL_SONG
        Dim cmdUpdSong
        Set cmdUpdSong = Server.CreateObject("ADODB.Command")
        cmdUpdSong.ActiveConnection = conn
        cmdUpdSong.CommandType = adCmdText
        cmdUpdSong.CommandText = "UPDATE TBL_SONG SET SONG_NAME = ?, SONG_PATH = ? WHERE ID_SONG = ?"
        cmdUpdSong.Parameters.Append cmdUpdSong.CreateParameter("@SONG_NAME", adVarChar, adParamInput, 50, editName)
        cmdUpdSong.Parameters.Append cmdUpdSong.CreateParameter("@SONG_PATH", adVarChar, adParamInput, 255, editPath)
        cmdUpdSong.Parameters.Append cmdUpdSong.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(editId))
        cmdUpdSong.Execute()

        ' Actualizar relación con Artista en TBL_PLAYER
        conn.Execute "DELETE FROM TBL_PLAYER WHERE ID_SONG = " & CLng(editId)
        If IsNumeric(editArtistId) And CLng(editArtistId) > 0 Then
            Dim cmdUpdArt
            Set cmdUpdArt = Server.CreateObject("ADODB.Command")
            cmdUpdArt.ActiveConnection = conn
            cmdUpdArt.CommandType = adCmdText
            cmdUpdArt.CommandText = "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (?, ?)"
            cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(editId))
            cmdUpdArt.Parameters.Append cmdUpdArt.CreateParameter("@ID_ARTIST", adInteger, adParamInput, , CLng(editArtistId))
            cmdUpdArt.Execute()
        End If

        ' Actualizar relación con Álbum en TBL_TRACKLIST
        conn.Execute "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = " & CLng(editId)
        If IsNumeric(editAlbumId) And CLng(editAlbumId) > 0 Then
            Dim cmdUpdAlb
            Set cmdUpdAlb = Server.CreateObject("ADODB.Command")
            cmdUpdAlb.ActiveConnection = conn
            cmdUpdAlb.CommandType = adCmdText
            cmdUpdAlb.CommandText = "INSERT INTO TBL_TRACKLIST (ID_ALBUM, ID_SONG) VALUES (?, ?)"
            cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@ID_ALBUM", adInteger, adParamInput, , CLng(editAlbumId))
            cmdUpdAlb.Parameters.Append cmdUpdAlb.CreateParameter("@ID_SONG", adInteger, adParamInput, , CLng(editId))
            cmdUpdAlb.Execute()
        End If

        infoMsg = "Canción actualizada exitosamente."
        On Error GoTo 0
    End If

' OPERACIÓN: ELIMINAR CANCIÓN
ElseIf action = "delete" Then
    Dim delId
    delId = Request("id")
    If IsNumeric(delId) Then
        On Error Resume Next
        ' 1. Eliminar de relaciones foráneas
        Dim cmdDelRel
        Set cmdDelRel = Server.CreateObject("ADODB.Command")
        cmdDelRel.ActiveConnection = conn
        cmdDelRel.CommandType = adCmdText
        
        cmdDelRel.CommandText = "DELETE FROM TBL_PLAYLIST_SONG WHERE ID_SONG = ?; " & _
                                "DELETE FROM TBL_PLAYER WHERE ID_SONG = ?; " & _
                                "DELETE FROM TBL_TRACKLIST WHERE ID_SONG = ?; " & _
                                "DELETE FROM TBL_SONG WHERE ID_SONG = ?;"
        cmdDelRel.Parameters.Append cmdDelRel.CreateParameter("@P1", adInteger, adParamInput, , CLng(delId))
        cmdDelRel.Parameters.Append cmdDelRel.CreateParameter("@P2", adInteger, adParamInput, , CLng(delId))
        cmdDelRel.Parameters.Append cmdDelRel.CreateParameter("@P3", adInteger, adParamInput, , CLng(delId))
        cmdDelRel.Parameters.Append cmdDelRel.CreateParameter("@P4", adInteger, adParamInput, , CLng(delId))
        cmdDelRel.Execute()

        infoMsg = "Canción eliminada del catálogo."
        On Error GoTo 0
    End If
End If

' ------------------------------------------------------------------------------
' 2. Consultas para Listado y Formularios
' ------------------------------------------------------------------------------
' Lista completa de canciones
Dim rsSongs
Set rsSongs = conn.Execute("SELECT " & _
    "s.ID_SONG, s.SONG_NAME, s.SONG_PATH, ISNULL(s.PLAYS, 0) AS PLAYS, " & _
    "(SELECT TOP 1 p.ID_ARTIST FROM TBL_PLAYER p WHERE p.ID_SONG = s.ID_SONG) AS ID_ARTIST, " & _
    "(SELECT TOP 1 art.ARTIST_NAME + ' ' + ISNULL(art.ARTIST_LASTNAME, '') " & _
    " FROM TBL_PLAYER p INNER JOIN TBL_ARTIST art ON p.ID_ARTIST = art.ID_ARTIST WHERE p.ID_SONG = s.ID_SONG) AS ARTIST_NAME, " & _
    "(SELECT TOP 1 tl.ID_ALBUM FROM TBL_TRACKLIST tl WHERE tl.ID_SONG = s.ID_SONG) AS ID_ALBUM, " & _
    "(SELECT TOP 1 alb.TITLE_ALBUM " & _
    " FROM TBL_TRACKLIST tl INNER JOIN TBL_ALBUM alb ON tl.ID_ALBUM = alb.ID_ALBUM WHERE tl.ID_SONG = s.ID_SONG) AS ALBUM_NAME " & _
    "FROM TBL_SONG s ORDER BY s.ID_SONG DESC")

' Lista de Artistas para desplegable
Dim rsArtists
Set rsArtists = conn.Execute("SELECT ID_ARTIST, ARTIST_NAME, ARTIST_LASTNAME FROM TBL_ARTIST ORDER BY ARTIST_NAME ASC")

' Lista de Álbumes para desplegable
Dim rsAlbums
Set rsAlbums = conn.Execute("SELECT ID_ALBUM, TITLE_ALBUM FROM TBL_ALBUM ORDER BY TITLE_ALBUM ASC")
%>
<!DOCTYPE html>
<html lang="es">
<head>
  <meta charset="UTF-8">
  <meta name="viewport" content="width=device-width, initial-scale=1.0">
  <title>CRUD Canciones • PoliMusic Admin</title>
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
            <h1 class="section-title">Gestión de Canciones (CRUD)</h1>
            <p style="color: var(--text-muted); font-size: 13.5px; margin-top: 4px;">Crear, editar, asociar con artistas/álbumes y eliminar pistas del catálogo</p>
          </div>
          <button type="button" class="btn btn-primary" onclick="openCreateModal()">
            <svg viewBox="0 0 24 24" width="16" height="16" fill="currentColor"><path d="M19 13h-6v6h-2v-6H5v-2h6V5h2v6h6v2z"/></svg>
            <span>Nueva Canción</span>
          </button>
        </div>

        <div class="admin-card">
          <div class="data-table-wrapper">
            <table class="data-table">
              <thead>
                <tr>
                  <th style="width: 50px;">ID</th>
                  <th>Título</th>
                  <th>Ruta de Archivo</th>
                  <th>Artista</th>
                  <th>Álbum</th>
                  <th style="text-align: right;">Plays</th>
                  <th style="width: 110px; text-align: center;">Acciones</th>
                </tr>
              </thead>
              <tbody>
                <%
                If Not rsSongs.EOF Then
                    Do While Not rsSongs.EOF
                        Dim artN, albN
                        artN = rsSongs("ARTIST_NAME")
                        If IsNull(artN) Or Trim(artN) = "" Then artN = "Sin asignar"
                        albN = rsSongs("ALBUM_NAME")
                        If IsNull(albN) Or Trim(albN) = "" Then albN = "Sin asignar"
                %>
                  <tr>
                    <td><strong>#<%=rsSongs("ID_SONG")%></strong></td>
                    <td style="color:#fff; font-weight: 600;"><%=SafeHtml(rsSongs("SONG_NAME"))%></td>
                    <td style="font-family: monospace; font-size: 12px; color: var(--text-dim);"><%=SafeHtml(rsSongs("SONG_PATH"))%></td>
                    <td><%=SafeHtml(artN)%></td>
                    <td><%=SafeHtml(albN)%></td>
                    <td style="text-align: right; font-variant-numeric: tabular-nums;"><%=rsSongs("PLAYS")%></td>
                    <td>
                      <div class="actions-cell" style="justify-content: center;">
                        <button type="button" class="btn-icon" title="Editar Canción" 
                                data-id="<%=rsSongs("ID_SONG")%>"
                                data-name="<%=SafeHtml(rsSongs("SONG_NAME"))%>"
                                data-path="<%=SafeHtml(rsSongs("SONG_PATH"))%>"
                                data-artist="<%=rsSongs("ID_ARTIST")%>"
                                data-album="<%=rsSongs("ID_ALBUM")%>"
                                onclick="handleEditSong(this)">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M3 17.25V21h3.75L17.81 9.94l-3.75-3.75L3 17.25zM20.71 7.04c.39-.39.39-1.02 0-1.41l-2.34-2.34c-.39-.39-1.02-.39-1.41 0l-1.83 1.83 3.75 3.75 1.83-1.83z"/></svg>
                        </button>
                        <button type="button" class="btn-icon delete" title="Eliminar Canción" 
                                data-id="<%=rsSongs("ID_SONG")%>"
                                data-name="<%=SafeHtml(rsSongs("SONG_NAME"))%>"
                                onclick="handleDeleteSong(this)">
                          <svg viewBox="0 0 24 24" width="15" height="15" fill="currentColor"><path d="M6 19c0 1.1.9 2 2 2h8c1.1 0 2-.9 2-2V7H6v12zM19 4h-3.5l-1-1h-5l-1 1H5v2h14V4z"/></svg>
                        </button>
                      </div>
                    </td>
                  </tr>
                <%
                        rsSongs.MoveNext
                    Loop
                Else
                    Response.Write "<tr><td colspan=""7"" style=""text-align:center; padding:24px;"">No hay canciones registradas.</td></tr>"
                End If
                CloseRecordset rsSongs
                %>
              </tbody>
            </table>
          </div>
        </div>
      </div>
    </main>
  </div>

  <!-- Modal Crear Canción -->
  <div id="modalCreateSong" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Nueva Canción</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalCreateSong')">&times;</button>
      </div>
      <form action="admin_songs.asp" method="POST">
        <input type="hidden" name="action" value="create">

        <div class="form-group">
          <label class="form-label" for="newSongName">Título de la Canción *</label>
          <input type="text" id="newSongName" name="song_name" class="form-control" placeholder="Ej. Bohemian Rhapsody" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="newSongPath">Ruta del Archivo de Audio *</label>
          <select id="newSongPath" name="song_path" class="form-control" required>
            <option value="~/songFiles/bensound-adventure.mp3">~/songFiles/bensound-adventure.mp3</option>
            <option value="~/songFiles/bensound-dreams.mp3">~/songFiles/bensound-dreams.mp3</option>
            <option value="~/songFiles/bensound-energy.mp3">~/songFiles/bensound-energy.mp3</option>
            <option value="~/songFiles/bensound-scifi.mp3">~/songFiles/bensound-scifi.mp3</option>
            <option value="~/songFiles/bensound-tomorrow.mp3">~/songFiles/bensound-tomorrow.mp3</option>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="newArtistId">Artista Intérprete</label>
          <select id="newArtistId" name="artist_id" class="form-control">
            <option value="0">-- Sin artista asignado --</option>
            <%
            If Not rsArtists.EOF Then
                Do While Not rsArtists.EOF
                    Response.Write "<option value=""" & rsArtists("ID_ARTIST") & """>" & SafeHtml(rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME")) & "</option>"
                    rsArtists.MoveNext
                Loop
            End If
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="newAlbumId">Álbum Discográfico</label>
          <select id="newAlbumId" name="album_id" class="form-control">
            <option value="0">-- Sencillo / Sin álbum --</option>
            <%
            If Not rsAlbums.EOF Then
                Do While Not rsAlbums.EOF
                    Response.Write "<option value=""" & rsAlbums("ID_ALBUM") & """>" & SafeHtml(rsAlbums("TITLE_ALBUM")) & "</option>"
                    rsAlbums.MoveNext
                Loop
            End If
            %>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalCreateSong')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Guardar Canción</button>
        </div>
      </form>
    </div>
  </div>

  <!-- Modal Editar Canción -->
  <div id="modalEditSong" class="modal-overlay">
    <div class="modal-content">
      <div class="modal-header">
        <h3 class="modal-title">Editar Canción</h3>
        <button type="button" class="modal-close" onclick="closeModal('modalEditSong')">&times;</button>
      </div>
      <form action="admin_songs.asp" method="POST">
        <input type="hidden" name="action" value="edit">
        <input type="hidden" id="editSongId" name="id_song" value="">

        <div class="form-group">
          <label class="form-label" for="editSongName">Título de la Canción *</label>
          <input type="text" id="editSongName" name="song_name" class="form-control" required maxlength="50">
        </div>

        <div class="form-group">
          <label class="form-label" for="editSongPath">Ruta del Archivo de Audio *</label>
          <input type="text" id="editSongPath" name="song_path" class="form-control" list="songFilesList" required maxlength="255">
          <datalist id="songFilesList">
            <option value="~/SongFiles/bensound-adventure.mp3">bensound-adventure.mp3 (Aventura / Rock Acústico)</option>
            <option value="~/SongFiles/bensound-dreams.mp3">bensound-dreams.mp3 (Chill / Piano Relajante)</option>
            <option value="~/SongFiles/bensound-energy.mp3">bensound-energy.mp3 (Energía / Pop Rock Activo)</option>
            <option value="~/SongFiles/bensound-scifi.mp3">bensound-scifi.mp3 (Electrónica / Synthwave Espacial)</option>
            <option value="~/SongFiles/bensound-tomorrow.mp3">bensound-tomorrow.mp3 (Inspiracional / Cinemático)</option>
          </datalist>
          <span style="font-size:11px; color: var(--text-dim); margin-top:4px; display:block;">Puedes seleccionar uno de los audios existentes en la carpeta <code>SongFiles/</code> o escribir la ruta de cualquier archivo MP3 nuevo.</span>
        </div>

        <div class="form-group">
          <label class="form-label" for="editArtistId">Artista Intérprete</label>
          <select id="editArtistId" name="artist_id" class="form-control">
            <option value="0">-- Sin artista asignado --</option>
            <%
            If rsArtists.State = 1 Then rsArtists.MoveFirst
            Do While Not rsArtists.EOF
                Response.Write "<option value=""" & rsArtists("ID_ARTIST") & """>" & SafeHtml(rsArtists("ARTIST_NAME") & " " & rsArtists("ARTIST_LASTNAME")) & "</option>"
                rsArtists.MoveNext
            Loop
            CloseRecordset rsArtists
            %>
          </select>
        </div>

        <div class="form-group">
          <label class="form-label" for="editAlbumId">Álbum Discográfico</label>
          <select id="editAlbumId" name="album_id" class="form-control">
            <option value="0">-- Sencillo / Sin álbum --</option>
            <%
            If rsAlbums.State = 1 Then rsAlbums.MoveFirst
            Do While Not rsAlbums.EOF
                Response.Write "<option value=""" & rsAlbums("ID_ALBUM") & """>" & SafeHtml(rsAlbums("TITLE_ALBUM")) & "</option>"
                rsAlbums.MoveNext
            Loop
            CloseRecordset rsAlbums
            %>
          </select>
        </div>

        <div class="modal-actions">
          <button type="button" class="btn btn-secondary" onclick="closeModal('modalEditSong')">Cancelar</button>
          <button type="submit" class="btn btn-primary">Actualizar Canción</button>
        </div>
      </form>
    </div>
  </div>

  <!--#include file="includes/player.asp"-->

  <script src="js/app.js"></script>
  <script src="js/player.js"></script>
  <script>
    function openCreateModal() {
      openModal('modalCreateSong');
    }

    function handleEditSong(btn) {
      document.getElementById('editSongId').value = btn.getAttribute('data-id');
      document.getElementById('editSongName').value = btn.getAttribute('data-name');
      document.getElementById('editSongPath').value = btn.getAttribute('data-path');
      document.getElementById('editArtistId').value = btn.getAttribute('data-artist') || '0';
      document.getElementById('editAlbumId').value = btn.getAttribute('data-album') || '0';
      openModal('modalEditSong');
    }

    function handleDeleteSong(btn) {
      var id = btn.getAttribute('data-id');
      var name = btn.getAttribute('data-name');
      confirmAction('¿Eliminar la canción \'' + name + '\' del catálogo?', 'admin_songs.asp?action=delete&id=' + id);
    }

    function openEditModal(id, name, path, artistId, albumId) {
      document.getElementById('editSongId').value = id;
      document.getElementById('editSongName').value = name;
      document.getElementById('editSongPath').value = path;
      document.getElementById('editArtistId').value = artistId || '0';
      document.getElementById('editAlbumId').value = albumId || '0';
      openModal('modalEditSong');
    }
  </script>
</body>
</html>
<%
CloseConnection conn
%>
