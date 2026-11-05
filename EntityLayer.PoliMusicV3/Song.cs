using System;

namespace EntityLayer.PoliMusicV3
{
    public class Song
    {
        // Propiedades de la tabla TBL_SONG
        public int Id { get; set; }
        public string Title { get; set; }
        public string FilePath { get; set; }
        public int Plays { get; set; }

        // Propiedad adicional para manejar la tabla intermedia TBL_PLAYER
        public int ArtistId { get; set; }

        // Propiedad opcional si en el futuro necesitas vincularla directo a un álbum desde la UI
        public int AlbumId { get; set; }
    }
}