/**
 * PoliMusic - Core Audio Player Engine (Spotify-Style)
 */

(function () {
  'use strict';

  // DOM Elements
  const audio = document.getElementById('poliAudio');
  if (!audio) return;

  const playBtn = document.getElementById('poliPlayBtn');
  const prevBtn = document.getElementById('poliPrevBtn');
  const nextBtn = document.getElementById('poliNextBtn');
  const shuffleBtn = document.getElementById('poliShuffleBtn');
  const repeatBtn = document.getElementById('poliRepeatBtn');

  const coverEl = document.getElementById('playerCover');
  const titleEl = document.getElementById('playerTitle');
  const artistEl = document.getElementById('playerArtist');

  const currentTimeEl = document.getElementById('currentTime');
  const totalDurationEl = document.getElementById('totalDuration');
  const progressBg = document.getElementById('progressBarBg');
  const progressFill = document.getElementById('progressBarFill');

  const volumeBg = document.getElementById('volumeBg');
  const volumeFill = document.getElementById('volumeFill');
  const volumeBtn = document.getElementById('volumeBtn');

  // Player State
  let currentTrack = null;
  let trackQueue = [];
  let queueIndex = -1;
  let isShuffle = false;
  let isRepeat = false;
  let lastVolume = 0.8;

  // Initialize volume
  audio.volume = 0.8;

  // Format Seconds to MM:SS
  function formatTime(seconds) {
    if (isNaN(seconds) || seconds < 0) return '0:00';
    const m = Math.floor(seconds / 60);
    const s = Math.floor(seconds % 60);
    return m + ':' + (s < 10 ? '0' : '') + s;
  }

  // Update UI with playing track
  function updateTrackUI(track) {
    if (!track) return;
    if (titleEl) titleEl.textContent = track.name || 'Canción sin título';
    if (artistEl) artistEl.textContent = track.artist || 'Artista PoliMusic';
    if (coverEl) {
      coverEl.src = track.cover || 'img/default-cover.svg';
      coverEl.alt = track.name || 'PoliMusic Cover';
    }
    document.title = (track.name ? track.name + ' • ' : '') + 'PoliMusic';
  }

  // Increment play counter in backend
  function incrementPlayCount(songId) {
    if (!songId) return;
    fetch('play_counter.asp?id=' + encodeURIComponent(songId), { method: 'POST' })
      .catch(function (err) {
        console.warn('Play counter notice:', err);
      });
  }

  // Play a specific track
  window.playTrack = function (id, name, artist, fileUrl, coverUrl) {
    if (!fileUrl) return;

    currentTrack = {
      id: id,
      name: name,
      artist: artist,
      url: fileUrl,
      cover: coverUrl
    };

    // Save in session storage
    try {
      sessionStorage.setItem('polimusic_last_track', JSON.stringify(currentTrack));
    } catch (e) {}

    updateTrackUI(currentTrack);

    audio.src = fileUrl;
    audio.load();
    const playPromise = audio.play();
    if (playPromise !== undefined) {
      playPromise
        .then(function () {
          updatePlayButton(true);
          incrementPlayCount(id);
        })
        .catch(function (err) {
          console.warn('Auto-play blocked or error:', err);
          updatePlayButton(false);
        });
    }
  };

  // Queue a list of songs
  window.playQueue = function (tracks, startIndex) {
    if (!tracks || !tracks.length) return;
    trackQueue = tracks;
    queueIndex = startIndex || 0;
    const t = trackQueue[queueIndex];
    if (t) {
      window.playTrack(t.id, t.name, t.artist, t.url, t.cover);
    }
  };

  // Toggle Play / Pause
  function togglePlay() {
    if (!audio.src || audio.src === window.location.href) {
      // If nothing loaded yet, try to play the first available track on the page
      const firstPlayBtn = document.querySelector('[data-audio-src]');
      if (firstPlayBtn) {
        firstPlayBtn.click();
        return;
      }
      return;
    }

    if (audio.paused) {
      audio.play().then(function () {
        updatePlayButton(true);
      });
    } else {
      audio.pause();
      updatePlayButton(false);
    }
  }

  function updatePlayButton(isPlaying) {
    if (!playBtn) return;
    const playIcon = playBtn.querySelector('.icon-play');
    const pauseIcon = playBtn.querySelector('.icon-pause');
    if (isPlaying) {
      if (playIcon) playIcon.style.display = 'none';
      if (pauseIcon) pauseIcon.style.display = 'block';
      playBtn.classList.add('playing');
    } else {
      if (playIcon) playIcon.style.display = 'block';
      if (pauseIcon) pauseIcon.style.display = 'none';
      playBtn.classList.remove('playing');
    }
  }

  // Audio Event Listeners
  audio.addEventListener('timeupdate', function () {
    if (!isNaN(audio.duration) && audio.duration > 0) {
      const percent = (audio.currentTime / audio.duration) * 100;
      if (progressFill) progressFill.style.width = percent + '%';
      if (currentTimeEl) currentTimeEl.textContent = formatTime(audio.currentTime);
    }
  });

  audio.addEventListener('loadedmetadata', function () {
    if (totalDurationEl && !isNaN(audio.duration)) {
      totalDurationEl.textContent = formatTime(audio.duration);
    }
  });

  audio.addEventListener('ended', function () {
    if (isRepeat) {
      audio.currentTime = 0;
      audio.play();
    } else if (trackQueue.length > 0 && queueIndex + 1 < trackQueue.length) {
      queueIndex++;
      const next = trackQueue[queueIndex];
      window.playTrack(next.id, next.name, next.artist, next.url, next.cover);
    } else {
      updatePlayButton(false);
    }
  });

  // Seeking on progress bar
  if (progressBg) {
    progressBg.addEventListener('click', function (e) {
      if (isNaN(audio.duration) || audio.duration <= 0) return;
      const rect = progressBg.getBoundingClientRect();
      const clickX = e.clientX - rect.left;
      const width = rect.width;
      const pct = Math.max(0, Math.min(1, clickX / width));
      audio.currentTime = pct * audio.duration;
    });
  }

  // Volume control
  if (volumeBg) {
    volumeBg.addEventListener('click', function (e) {
      const rect = volumeBg.getBoundingClientRect();
      const clickX = e.clientX - rect.left;
      const pct = Math.max(0, Math.min(1, clickX / rect.width));
      audio.volume = pct;
      if (volumeFill) volumeFill.style.width = (pct * 100) + '%';
      if (pct > 0) lastVolume = pct;
    });
  }

  if (volumeBtn) {
    volumeBtn.addEventListener('click', function () {
      if (audio.volume > 0) {
        lastVolume = audio.volume;
        audio.volume = 0;
        if (volumeFill) volumeFill.style.width = '0%';
      } else {
        audio.volume = lastVolume || 0.8;
        if (volumeFill) volumeFill.style.width = ((lastVolume || 0.8) * 100) + '%';
      }
    });
  }

  // Next and Prev handlers
  function nextTrack() {
    if (trackQueue.length > 0) {
      if (isShuffle) {
        queueIndex = Math.floor(Math.random() * trackQueue.length);
      } else {
        queueIndex = (queueIndex + 1) % trackQueue.length;
      }
      const t = trackQueue[queueIndex];
      window.playTrack(t.id, t.name, t.artist, t.url, t.cover);
    } else {
      // Seek slightly forward
      audio.currentTime = Math.min(audio.duration || 0, audio.currentTime + 10);
    }
  }

  function prevTrack() {
    if (audio.currentTime > 3) {
      audio.currentTime = 0;
      return;
    }
    if (trackQueue.length > 0) {
      queueIndex = (queueIndex - 1 + trackQueue.length) % trackQueue.length;
      const t = trackQueue[queueIndex];
      window.playTrack(t.id, t.name, t.artist, t.url, t.cover);
    } else {
      audio.currentTime = 0;
    }
  }

  if (playBtn) playBtn.addEventListener('click', togglePlay);
  if (nextBtn) nextBtn.addEventListener('click', nextTrack);
  if (prevBtn) prevBtn.addEventListener('click', prevTrack);

  if (repeatBtn) {
    repeatBtn.addEventListener('click', function () {
      isRepeat = !isRepeat;
      repeatBtn.classList.toggle('active', isRepeat);
    });
  }

  if (shuffleBtn) {
    shuffleBtn.addEventListener('click', function () {
      isShuffle = !isShuffle;
      shuffleBtn.classList.toggle('active', isShuffle);
    });
  }

  // Bind all track play buttons across any page
  document.addEventListener('click', function (e) {
    const trigger = e.target.closest('[data-audio-src]');
    if (!trigger) return;

    e.preventDefault();
    const id = trigger.getAttribute('data-song-id');
    const name = trigger.getAttribute('data-song-name');
    const artist = trigger.getAttribute('data-artist-name');
    const src = trigger.getAttribute('data-audio-src');
    const cover = trigger.getAttribute('data-cover-src');

    // Build queue from all songs available in the current context/table
    const allSongTriggers = Array.from(document.querySelectorAll('[data-audio-src]'));
    if (allSongTriggers.length > 1) {
      const q = allSongTriggers.map(function (el) {
        return {
          id: el.getAttribute('data-song-id'),
          name: el.getAttribute('data-song-name'),
          artist: el.getAttribute('data-artist-name'),
          url: el.getAttribute('data-audio-src'),
          cover: el.getAttribute('data-cover-src')
        };
      });
      const idx = allSongTriggers.indexOf(trigger);
      window.playQueue(q, idx >= 0 ? idx : 0);
    } else {
      window.playTrack(id, name, artist, src, cover);
    }
  });

  // Restore previous track state if available
  try {
    const saved = sessionStorage.getItem('polimusic_last_track');
    if (saved) {
      const t = JSON.parse(saved);
      if (t && t.url) {
        currentTrack = t;
        updateTrackUI(t);
        audio.src = t.url;
      }
    }
  } catch (e) {}
})();
