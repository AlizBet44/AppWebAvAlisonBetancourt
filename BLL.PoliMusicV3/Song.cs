using System;
using System.Collections.Generic;
using EntityLayer.PoliMusicV3;

namespace BLL.PoliMusicV3
{
    public class Song
    {
        private string _connectionString;

        public Song(string connectionString)
        {
            _connectionString = connectionString;
        }

        public List<EntityLayer.PoliMusicV3.Song> GetAllSongs()
        {
            DAL.PoliMusicV3.Song dalSong = new DAL.PoliMusicV3.Song(_connectionString);
            return dalSong.GetAll();
        }

        public string Save(EntityLayer.PoliMusicV3.Song song)
        {
            if (string.IsNullOrWhiteSpace(song.Title))
            {
                return "El nombre de la canción es obligatorio.";
            }

            if (song.ArtistId <= 0)
            {
                return "Debe asignar un ID de artista válido existente en TBL_ARTIST.";
            }

            DAL.PoliMusicV3.Song dalSong = new DAL.PoliMusicV3.Song(_connectionString);
            return dalSong.Insert(song);
        }
    }
}