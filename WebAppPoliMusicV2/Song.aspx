<%@ Page Language="C#" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV2.Song" %>

<!DOCTYPE html>
<html xmlns="http://www.w3.org/1999/xhtml">
<head runat="server">
    <title>Songs - PoliMusic</title>
    <style>
        body {
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            background-color: #fcfcfc;
            color: #111111;
            margin: 0;
            padding: 40px 20px;
            display: flex;
            justify-content: center;
        }
        .container {
            width: 100%;
            max-width: 950px;
            background: #ffffff;
            padding: 40px;
            border-radius: 20px;
            box-shadow: 0 10px 40px rgba(0,0,0,0.06);
            border-top: 5px solid #9c27b0;
        }
        .header-top {
            display: flex;
            justify-content: space-between;
            align-items: center;
            margin-bottom: 30px;
            padding-bottom: 20px;
            border-bottom: 1px solid #f3e5f5;
        }
        h2 {
            color: #4a148c;
            margin: 0;
            font-weight: 700;
            font-size: 26px;
        }
        .logout-link {
            background: linear-gradient(135deg, #e91e63, #9c27b0);
            color: #ffffff;
            text-decoration: none;
            padding: 10px 25px;
            border-radius: 30px;
            font-weight: 600;
            font-size: 14px;
            transition: transform 0.2s, box-shadow 0.2s;
            box-shadow: 0 4px 15px rgba(156, 39, 176, 0.3);
        }
        .logout-link:hover {
            transform: translateY(-2px);
            box-shadow: 0 6px 20px rgba(156, 39, 176, 0.4);
            color: #ffffff;
        }
        .subtitle {
            color: #333333;
            font-weight: 600;
            font-size: 18px;
            margin-bottom: 15px;
            text-align: left;
        }
        .grid-view {
            width: 100%;
            border-collapse: collapse;
            text-align: center;
        }
        .grid-view th {
            background-color: #ffffff;
            color: #000000;
            padding: 15px;
            border-bottom: 2px solid #ce93d8;
            font-weight: 600;
            text-transform: uppercase;
            font-size: 13px;
            letter-spacing: 1px;
        }
        .grid-view td {
            padding: 15px;
            border-bottom: 1px solid #f3e5f5;
            vertical-align: middle;
            color: #333333;
        }
        .grid-view tr:hover {
            background-color: #fbf5ff;
        }
        audio {
            height: 35px;
            outline: none;
        }
    </style>
</head>
<body>
    <form id="form1" runat="server">
        <div class="container">
            <div class="header-top">
                <h2>Welcome <asp:Label ID="lblUserName" runat="server" Text="Alison.Betancourt"></asp:Label> </h2>
                <asp:LinkButton ID="lnkLogout" runat="server" CssClass="logout-link" OnClick="lnkLogout_Click">Cerrar Sesión</asp:LinkButton>
            </div>
            
            <div class="subtitle">Song list:</div>
            
            <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" CssClass="grid-view" GridLines="None">
                <Columns>
                    <asp:BoundField DataField="ID_SONG" HeaderText="ID" />
                    <asp:BoundField DataField="SONG_NAME" HeaderText="Nombre" />
                    <asp:BoundField DataField="SONG_PATH" HeaderText="Ruta" />
                    <asp:BoundField DataField="PLAYS" HeaderText="Plays" />
                    <asp:TemplateField HeaderText="Reproducir">
                        <ItemTemplate>
                            <audio controls="controls">
                                <source src='<%# ResolveUrl(Eval("SONG_PATH").ToString()) %>' type="audio/mpeg" />
                            </audio>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </div>
    </form>
</body>
</html>