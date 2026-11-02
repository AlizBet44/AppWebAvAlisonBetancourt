using System;
using System.Configuration;
using System.Data;

namespace WebAppPoliMusicV2
{
    public partial class Song : System.Web.UI.Page
    {
        protected void Page_Load(object sender, EventArgs e)
        {
            if (!IsPostBack)
            {
                // Extrae el nombre del usuario de la sesión para el saludo inicial
                if (Session[Util.Constants.USER] != null)
                {
                    Entities.User user = (Entities.User)Session[Util.Constants.USER];
                    lblUserName.Text = user.UserName;
                }

                LoadSongs();
            }
        }
        private void LoadSongs()
        {
            string strConnString = ConfigurationManager.ConnectionStrings["BDD_PoliMusicConnectionString"].ConnectionString;
            BLL.Song songBLL = new BLL.Song(strConnString);

            DataTable dtSongs = songBLL.Read();

            gvSongs.DataSource = dtSongs;
            gvSongs.DataBind();
        }

        protected void lnkLogout_Click(object sender, EventArgs e)
        {
            Response.Redirect("Logout.aspx");
        }
    }
}