/**
 * PoliMusic - Global Application UI Scripts
 */

// Modal System Helpers
function openModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.add('active');
    document.body.style.overflow = 'hidden';
  }
}

function closeModal(modalId) {
  const modal = document.getElementById(modalId);
  if (modal) {
    modal.classList.remove('active');
    document.body.style.overflow = '';
  }
}

// Confirmation helper for CRUD deletes
function confirmAction(message, url) {
  if (confirm(message || '¿Está seguro de realizar esta acción?')) {
    window.location.href = url;
  }
}

// Client-side quick filter for songs/tables/cards
document.addEventListener('DOMContentLoaded', function () {
  const searchInput = document.getElementById('globalSearchInput');
  if (searchInput) {
    searchInput.addEventListener('input', function () {
      const q = searchInput.value.toLowerCase().trim();

      // Filter song table rows
      const songRows = document.querySelectorAll('.song-row');
      songRows.forEach(function (row) {
        const text = row.textContent.toLowerCase();
        row.style.display = text.indexOf(q) !== -1 ? '' : 'none';
      });

      // Filter cards
      const cards = document.querySelectorAll('.music-card, .artist-card');
      cards.forEach(function (card) {
        const text = card.textContent.toLowerCase();
        card.style.display = text.indexOf(q) !== -1 ? '' : 'none';
      });
    });
  }

  // Close modals when clicking overlay background
  document.querySelectorAll('.modal-overlay').forEach(function (overlay) {
    overlay.addEventListener('click', function (e) {
      if (e.target === overlay) {
        overlay.classList.remove('active');
        document.body.style.overflow = '';
      }
    });
  });
});
