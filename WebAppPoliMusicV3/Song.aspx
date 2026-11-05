<%@ Page Title="Canciones" Language="C#" MasterPageFile="~/Site.Master" AutoEventWireup="true" CodeBehind="Song.aspx.cs" Inherits="WebAppPoliMusicV3.Song" %>

<asp:Content ID="Content1" ContentPlaceHolderID="MainContent" runat="server">
    <div style="background-color: #FFFFFF; border: 2px solid #F48FB1; border-radius: 12px; padding: 25px; margin-top: 20px; box-shadow: 0 4px 8px rgba(0,0,0,0.05);">
        
        <h2 style="color: #4A148C; font-weight: bold; margin-bottom: 20px; font-family: Arial, sans-serif;">
            Song list:
        </h2>

        <div style="overflow-x: auto;">
            <asp:GridView ID="gvSongs" runat="server" AutoGenerateColumns="False" 
                          Width="100%" GridLines="Both" 
                          BorderColor="#F48FB1" BorderStyle="Solid" BorderWidth="2px"
                          BackColor="#FFFFFF" CellPadding="10" 
                          Font-Names="Arial, sans-serif">
                
                <HeaderStyle BackColor="#9C27B0" ForeColor="#FFFFFF" Font-Bold="True" HorizontalAlign="Center" Font-Size="16px" />
                <RowStyle BackColor="#FFFFFF" ForeColor="#000000" Font-Size="14px" />
                <AlternatingRowStyle BackColor="#FCE4EC" /> 
                
                <Columns>
                    <asp:BoundField DataField="Id" HeaderText="ID Canción" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="10%" />
                    <asp:BoundField DataField="Title" HeaderText="Nombre" ItemStyle-Width="20%" />
                    <asp:BoundField DataField="FilePath" HeaderText="Ruta" ItemStyle-Width="30%" />
                    <asp:BoundField DataField="Plays" HeaderText="Reproducciones" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="15%" />
                    
                    <asp:TemplateField HeaderText="Reproducir" ItemStyle-HorizontalAlign="Center" ItemStyle-Width="25%">
                        <ItemTemplate>
                            <audio controls style="height: 35px; width: 100%; outline: none;">
                                <source src='<%# ResolveUrl(Eval("FilePath").ToString()) %>' type="audio/mpeg">
                                Tu navegador no soporta el elemento de audio.
                            </audio>
                        </ItemTemplate>
                    </asp:TemplateField>
                </Columns>
            </asp:GridView>
        </div>
    </div>
</asp:Content>