using System.Data;

namespace WebAppPoliMusicV2.BLL
{
    public class Song
    {
        private string strConnectionString;

        public Song(string strConnString)
        {
            strConnectionString = strConnString;
        }

        public DataTable Read()
        {
            return new DAL.Song(strConnectionString).Read();
        }
    }
}