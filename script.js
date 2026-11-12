/* ==========================================================================
   마케터 장상균 포트폴리오 — Green Grid Template Interactions (script.js)
   ========================================================================== */

document.addEventListener('DOMContentLoaded', () => {
  const urlParams = new URLSearchParams(window.location.search);
  const captureParam = urlParams.get('capture');
  if (captureParam) {
    document.body.classList.add('capture-slide-mode');
    const targetSheet = document.getElementById(`slide-${captureParam}`);
    if (targetSheet) {
      targetSheet.classList.add('active-capture');
    }
    return;
  }

  const sheets = Array.from(document.querySelectorAll('.template-sheet'));
  const navLinks = Array.from(document.querySelectorAll('.quick-nav .nav-link'));
  const indexEntries = Array.from(document.querySelectorAll('.index-entry'));

  const btnScrollMode = document.getElementById('btnScrollMode');
  const btnDeckMode = document.getElementById('btnDeckMode');
  const btnPrintPdf = document.getElementById('btnPrintPdf');

  const prevSlideBtn = document.getElementById('prevSlideBtn');
  const nextSlideBtn = document.getElementById('nextSlideBtn');
  const slideProgressText = document.getElementById('slideProgressText');

  const lightboxModal = document.getElementById('lightboxModal');
  const lightboxImg = document.getElementById('lightboxImg');
  const lightboxClose = document.getElementById('lightboxClose');

  let isDeckMode = false;
  let currentSlideIndex = 0;

  /* --------------------------------------------------------------------------
     1. Slide Deck Mode vs. Web Scroll Mode
     -------------------------------------------------------------------------- */
  function updateDeckUI() {
    sheets.forEach((sheet, idx) => {
      sheet.classList.toggle('active-slide', idx === currentSlideIndex);
    });

    const formattedCurrent = String(currentSlideIndex + 1).padStart(2, '0');
    const formattedTotal = String(sheets.length).padStart(2, '0');
    if (slideProgressText) {
      slideProgressText.textContent = `${formattedCurrent} / ${formattedTotal}`;
    }

    if (prevSlideBtn) prevSlideBtn.disabled = currentSlideIndex === 0;
    if (nextSlideBtn) nextSlideBtn.disabled = currentSlideIndex === sheets.length - 1;

    updateActiveNavLink(currentSlideIndex);
  }

  function updateActiveNavLink(slideIdx) {
    let bestMatch = navLinks[0];
    navLinks.forEach((link) => {
      const targetIdx = parseInt(link.getAttribute('data-slide'), 10);
      if (targetIdx <= slideIdx) {
        bestMatch = link;
      }
    });
    navLinks.forEach((link) => link.classList.toggle('active', link === bestMatch));
  }

  function setMode(deckMode) {
    isDeckMode = deckMode;
    document.body.classList.toggle('deck-mode', isDeckMode);
    if (btnScrollMode) btnScrollMode.classList.toggle('active', !isDeckMode);
    if (btnDeckMode) btnDeckMode.classList.toggle('active', isDeckMode);

    if (isDeckMode) {
      updateDeckUI();
    } else {
      sheets.forEach((sheet) => sheet.classList.remove('active-slide'));
      if (sheets[currentSlideIndex]) {
        sheets[currentSlideIndex].scrollIntoView({ behavior: 'smooth', block: 'start' });
      }
    }
  }

  function goToSlide(index) {
    const clamped = Math.max(0, Math.min(sheets.length - 1, index));
    currentSlideIndex = clamped;

    if (isDeckMode) {
      updateDeckUI();
    } else if (sheets[currentSlideIndex]) {
      sheets[currentSlideIndex].scrollIntoView({ behavior: 'smooth', block: 'start' });
      updateActiveNavLink(currentSlideIndex);
    }
  }

  if (btnScrollMode) {
    btnScrollMode.addEventListener('click', () => setMode(false));
  }
  if (btnDeckMode) {
    btnDeckMode.addEventListener('click', () => setMode(true));
  }
  if (btnPrintPdf) {
    btnPrintPdf.addEventListener('click', () => {
      const wasDeck = isDeckMode;
      if (wasDeck) document.body.classList.remove('deck-mode');
      window.print();
      if (wasDeck) document.body.classList.add('deck-mode');
    });
  }

  if (prevSlideBtn) {
    prevSlideBtn.addEventListener('click', () => goToSlide(currentSlideIndex - 1));
  }
  if (nextSlideBtn) {
    nextSlideBtn.addEventListener('click', () => goToSlide(currentSlideIndex + 1));
  }

  /* --------------------------------------------------------------------------
     2. Top Navigation & INDEX Slide Clicks
     -------------------------------------------------------------------------- */
  navLinks.forEach((link) => {
    link.addEventListener('click', (e) => {
      e.preventDefault();
      const targetSlide = parseInt(link.getAttribute('data-slide'), 10);
      if (!Number.isNaN(targetSlide)) {
        goToSlide(targetSlide);
      }
    });
  });

  indexEntries.forEach((entry) => {
    entry.addEventListener('click', (e) => {
      e.preventDefault();
      const targetSlide = parseInt(entry.getAttribute('data-target-slide'), 10);
      if (!Number.isNaN(targetSlide)) {
        goToSlide(targetSlide);
      }
    });
  });

  /* --------------------------------------------------------------------------
     3. Keyboard Navigation (Arrow Keys / PageUp / PageDown / Escape)
     -------------------------------------------------------------------------- */
  document.addEventListener('keydown', (e) => {
    if (lightboxModal && lightboxModal.classList.contains('open')) {
      if (e.key === 'Escape') {
        lightboxModal.classList.remove('open');
      }
      return;
    }

    if (!isDeckMode) return;

    if (e.key === 'ArrowRight' || e.key === 'PageDown' || e.key === ' ') {
      e.preventDefault();
      goToSlide(currentSlideIndex + 1);
    } else if (e.key === 'ArrowLeft' || e.key === 'PageUp') {
      e.preventDefault();
      goToSlide(currentSlideIndex - 1);
    } else if (e.key === 'Home') {
      e.preventDefault();
      goToSlide(0);
    } else if (e.key === 'End') {
      e.preventDefault();
      goToSlide(sheets.length - 1);
    }
  });

  /* --------------------------------------------------------------------------
     4. ScrollSpy in Web Scroll Mode
     -------------------------------------------------------------------------- */
  const observer = new IntersectionObserver(
    (entries) => {
      if (isDeckMode) return;
      entries.forEach((entry) => {
        if (entry.isIntersecting) {
          const idx = parseInt(entry.target.getAttribute('data-index'), 10);
          if (!Number.isNaN(idx)) {
            currentSlideIndex = idx;
            updateActiveNavLink(idx);
          }
        }
      });
    },
    { threshold: 0.45 }
  );

  sheets.forEach((sheet) => observer.observe(sheet));

  /* --------------------------------------------------------------------------
     5. Image Lightbox Zoom (Poster, Smartbin, Cherry Coffee, Profile)
     -------------------------------------------------------------------------- */
  const zoomableImages = document.querySelectorAll('.zoomable-img');
  zoomableImages.forEach((img) => {
    img.addEventListener('click', () => {
      if (!lightboxModal || !lightboxImg) return;
      lightboxImg.src = img.src;
      lightboxImg.alt = img.alt || '확대 이미지';
      lightboxModal.classList.add('open');
    });
  });

  if (lightboxClose) {
    lightboxClose.addEventListener('click', () => lightboxModal.classList.remove('open'));
  }
  if (lightboxModal) {
    lightboxModal.addEventListener('click', (e) => {
      if (e.target === lightboxModal) {
        lightboxModal.classList.remove('open');
      }
    });
  }

  /* --------------------------------------------------------------------------
     6. Copy-to-Clipboard Helper for Contact Info
     -------------------------------------------------------------------------- */
  function showToast(message) {
    let toast = document.getElementById('copyToast');
    if (!toast) {
      toast = document.createElement('div');
      toast.id = 'copyToast';
      Object.assign(toast.style, {
        position: 'fixed',
        bottom: '24px',
        right: '24px',
        background: '#268C3A',
        color: '#ffffff',
        padding: '10px 18px',
        fontWeight: '700',
        fontSize: '0.86rem',
        border: '2px solid #ffffff',
        boxShadow: '0 8px 24px rgba(0,0,0,0.22)',
        zIndex: '3000',
        transition: 'opacity 0.2s ease'
      });
      document.body.appendChild(toast);
    }
    toast.textContent = message;
    toast.style.opacity = '1';
    clearTimeout(toast._timer);
    toast._timer = setTimeout(() => {
      toast.style.opacity = '0';
    }, 2200);
  }

  document.querySelectorAll('[data-copy]').forEach((el) => {
    el.addEventListener('click', (e) => {
      const text = el.getAttribute('data-copy');
      if (text && navigator.clipboard) {
        e.preventDefault();
        navigator.clipboard.writeText(text).then(() => {
          showToast(`클립보드에 복사되었습니다: ${text}`);
        });
      }
    });
  });
});
