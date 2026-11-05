using System;
using System.Configuration;

namespace WebAppPoliMusicV3
{
    public partial class Song : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                LoadSongsGrid();
            }
        }

        private void LoadSongsGrid()
        {
            // Obtener la cadena de conexión desde el Web.config
            string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;

            // Instanciar la capa de negocio BLL
            BLL.PoliMusicV3.Song bllSong = new BLL.PoliMusicV3.Song(strConnString);

            // Enlazar la lista de canciones al GridView
            gvSongs.DataSource = bllSong.GetAllSongs();
            gvSongs.DataBind();
        }
    }
}