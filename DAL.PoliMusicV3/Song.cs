using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using EntityLayer.PoliMusicV3;

namespace DAL.PoliMusicV3
{
    public class Song
    {
        private string _connectionString;

        public Song(string connectionString)
        {
            _connectionString = connectionString;
        }

        public List<EntityLayer.PoliMusicV3.Song> GetAll()
        {
            List<EntityLayer.PoliMusicV3.Song> songs = new List<EntityLayer.PoliMusicV3.Song>();
            using (SqlConnection conn = new SqlConnection(_connectionString))
            {
                // Se consulta la tabla TBL_SONG según tu esquema
                string query = "SELECT ID_SONG, SONG_NAME, SONG_PATH, PLAYS FROM TBL_SONG";
                SqlCommand cmd = new SqlCommand(query, conn);
                conn.Open();

                using (SqlDataReader reader = cmd.ExecuteReader())
                {
                    while (reader.Read())
                    {
                        songs.Add(new EntityLayer.PoliMusicV3.Song
                        {
                            Id = Convert.ToInt32(reader["ID_SONG"]),
                            Title = reader["SONG_NAME"].ToString(),
                            FilePath = reader["SONG_PATH"].ToString()
                            // Nota: Asumimos que la entidad Song tiene estas propiedades. 
                            // Si tu entidad tiene 'ArtistId', se omite aquí porque requiere un JOIN con TBL_PLAYER.
                        });
                    }
                }
            }
            return songs;
        }

        public string Insert(EntityLayer.PoliMusicV3.Song song)
        {
            try
            {
                using (SqlConnection conn = new SqlConnection(_connectionString))
                {
                    conn.Open();
                    // Se utiliza una transacción para insertar en TBL_SONG y TBL_PLAYER de forma segura
                    using (SqlTransaction transaction = conn.BeginTransaction())
                    {
                        try
                        {
                            // 1. Insertar en TBL_SONG y obtener el ID generado[cite: 1]
                            string querySong = "INSERT INTO TBL_SONG (SONG_NAME, SONG_PATH, PLAYS) OUTPUT INSERTED.ID_SONG VALUES (@SongName, @SongPath, @Plays)";
                            SqlCommand cmdSong = new SqlCommand(querySong, conn, transaction);
                            cmdSong.Parameters.AddWithValue("@SongName", song.Title);
                            cmdSong.Parameters.AddWithValue("@SongPath", song.FilePath ?? (object)DBNull.Value);
                            cmdSong.Parameters.AddWithValue("@Plays", 0); // Inicializa con 0 reproducciones[cite: 1]

                            int insertedSongId = (int)cmdSong.ExecuteScalar();

                            // 2. Insertar en TBL_PLAYER la relación entre ID_SONG e ID_ARTIST[cite: 1]
                            string queryPlayer = "INSERT INTO TBL_PLAYER (ID_SONG, ID_ARTIST) VALUES (@IdSong, @IdArtist)";
                            SqlCommand cmdPlayer = new SqlCommand(queryPlayer, conn, transaction);
                            cmdPlayer.Parameters.AddWithValue("@IdSong", insertedSongId);
                            cmdPlayer.Parameters.AddWithValue("@IdArtist", song.ArtistId); // Asegúrate que EntityLayer.Song tenga la propiedad ArtistId

                            cmdPlayer.ExecuteNonQuery();

                            // Confirmar la transacción
                            transaction.Commit();
                            return string.Empty;
                        }
                        catch (Exception ex)
                        {
                            transaction.Rollback();
                            return "Error en la transacción SQL: " + ex.Message;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                return ex.Message;
            }
        }
    }
}