<%
' ==============================================================================
' PoliMusic - Reproductor de Audio Persistente Inferior (Spotify Experience)
' ==============================================================================
%>
<!-- Elemento de Audio Nativo HTML5 -->
<audio id="poliAudio" preload="metadata"></audio>

<!-- Barra Inferior del Reproductor -->
<footer class="poli-player-bar">
  <!-- 1. Información de la pista actual (Izquierda) -->
  <div class="player-track-info">
    <img id="playerCover" src="img/default-cover.svg" alt="PoliMusic Cover" class="player-cover">
    <div class="player-text">
      <span id="playerTitle" class="player-title">Selecciona una canción</span>
      <span id="playerArtist" class="player-artist">PoliMusic Player</span>
    </div>
  </div>

  <!-- 2. Controles de Reproducción y Línea de Tiempo (Centro) -->
  <div class="player-center">
    <div class="player-buttons">
      <!-- Botón Aleatorio -->
      <button type="button" id="poliShuffleBtn" class="btn-ctrl" title="Aleatorio">
        <svg viewBox="0 0 24 24"><path d="M10.59 9.17L5.41 4 4 5.41l5.17 5.17 1.42-1.41zM14.5 4l2.04 2.04L4 18.59 5.41 20 17.96 7.46 20 9.5V4h-5.5zm.33 9.41l-1.41 1.41 3.13 3.13L14.5 20H20v-5.5l-2.04 2.04-3.13-3.13z"/></svg>
      </button>

      <!-- Canción Anterior -->
      <button type="button" id="poliPrevBtn" class="btn-ctrl" title="Anterior">
        <svg viewBox="0 0 24 24"><path d="M6 6h2v12H6zm3.5 6l8.5 6V6z"/></svg>
      </button>

      <!-- Botón Principal Play / Pause -->
      <button type="button" id="poliPlayBtn" class="btn-play-main" title="Reproducir / Pausar">
        <svg class="icon-play" viewBox="0 0 24 24"><path d="M8 5v14l11-7z"/></svg>
        <svg class="icon-pause" viewBox="0 0 24 24" style="display:none;"><path d="M6 19h4V5H6v14zm8-14v14h4V5h-4z"/></svg>
      </button>

      <!-- Canción Siguiente -->
      <button type="button" id="poliNextBtn" class="btn-ctrl" title="Siguiente">
        <svg viewBox="0 0 24 24"><path d="M6 18l8.5-6L6 6v12zM16 6v12h2V6h-2z"/></svg>
      </button>

      <!-- Repetir Pista -->
      <button type="button" id="poliRepeatBtn" class="btn-ctrl" title="Repetir">
        <svg viewBox="0 0 24 24"><path d="M7 7h10v3l4-4-4-4v3H5v6h2V7zm10 10H7v-3l-4 4 4 4v-3h12v-6h-2v4z"/></svg>
      </button>
    </div>

    <!-- Barra de Progreso y Tiempo -->
    <div class="player-progress-container">
      <span id="currentTime" class="time-stamp">0:00</span>
      <div id="progressBarBg" class="progress-bar-bg" title="Adelantar / Retroceder">
        <div id="progressBarFill" class="progress-bar-fill">
          <div class="progress-thumb"></div>
        </div>
      </div>
      <span id="totalDuration" class="time-stamp">0:00</span>
    </div>
  </div>

  <!-- 3. Controles de Volumen (Derecha) -->
  <div class="player-right">
    <div class="volume-container">
      <button type="button" id="volumeBtn" class="btn-ctrl" title="Silenciar / Activar sonido">
        <svg viewBox="0 0 24 24"><path d="M3 9v6h4l5 5V4L7 9H3zm13.5 3c0-1.77-1.02-3.29-2.5-4.03v8.05c1.48-.73 2.5-2.25 2.5-4.02zM14 3.23v2.06c2.89.86 5 3.54 5 6.71s-2.11 5.85-5 6.71v2.06c4.01-.91 7-4.49 7-8.77s-2.99-7.86-7-8.77z"/></svg>
      </button>
      <div id="volumeBg" class="volume-slider-bg" title="Ajustar volumen">
        <div id="volumeFill" class="volume-slider-fill" style="width: 80%;"></div>
      </div>
    </div>
  </div>
</footer>
