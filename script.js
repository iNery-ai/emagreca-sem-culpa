document.addEventListener('DOMContentLoaded', () => {
  
  // ==========================================
  // 1. STICKY BAR VISIBILITY ON SCROLL
  // ==========================================
  const stickyBar = document.getElementById('stickyBar');
  const mobileStickyBtn = document.getElementById('mobileStickyBtn');
  const heroSection = document.querySelector('.hero');

  const handleScroll = () => {
    if (!heroSection) return;
    
    const heroHeight = heroSection.offsetHeight;
    const shouldBeVisible = window.scrollY > (heroHeight - 100);
    
    if (stickyBar) {
      if (shouldBeVisible) {
        stickyBar.classList.add('visible');
      } else {
        stickyBar.classList.remove('visible');
      }
    }
    
    if (mobileStickyBtn) {
      if (shouldBeVisible) {
        mobileStickyBtn.classList.add('visible');
      } else {
        mobileStickyBtn.classList.remove('visible');
      }
    }
  };

  window.addEventListener('scroll', handleScroll);
  // Run once initially in case the page loaded scrolled down
  handleScroll();


  // ==========================================
  // 2. SCROLL REVEAL (INTERSECTION OBSERVER)
  // ==========================================
  const revealElements = document.querySelectorAll('.scroll-reveal');
  
  if ('IntersectionObserver' in window) {
    const revealObserver = new IntersectionObserver((entries, observer) => {
      entries.forEach(entry => {
        if (entry.isIntersecting) {
          entry.target.classList.add('revealed');
          // Once animated, we don't need to observe it anymore
          observer.unobserve(entry.target);
        }
      });
    }, {
      threshold: 0.05, // trigger when 5% of the element is visible
      rootMargin: '0px 0px -50px 0px' // offset bottom trigger slightly
    });

    revealElements.forEach(el => revealObserver.observe(el));
  } else {
    // Fallback: Reveal all immediately if browser doesn't support IntersectionObserver
    revealElements.forEach(el => el.classList.add('revealed'));
  }


  // ==========================================
  // 3. FAQ ACCORDION LOGIC
  // ==========================================
  const faqQuestions = document.querySelectorAll('.faq-question');

  faqQuestions.forEach(question => {
    question.addEventListener('click', () => {
      const faqItem = question.parentElement;
      const faqAnswer = faqItem.querySelector('.faq-answer');
      const isActive = faqItem.classList.contains('active');

      // Close all other FAQ items first and reset their aria-expanded
      document.querySelectorAll('.faq-item').forEach(item => {
        item.classList.remove('active');
        item.querySelector('.faq-answer').style.maxHeight = null;
        item.querySelector('.faq-question').setAttribute('aria-expanded', 'false');
      });

      // Toggle current item
      if (!isActive) {
        faqItem.classList.add('active');
        // Set max-height to the exact content scrollHeight for CSS animation
        faqAnswer.style.maxHeight = faqAnswer.scrollHeight + 'px';
        // Mark the clicked question as expanded for accessibility
        question.setAttribute('aria-expanded', 'true');
      }
    });
  });


  // ==========================================
  // 4. EVERGREEN COUNTDOWN TIMER (48 HOURS)
  // ==========================================
  const COUNTDOWN_STORAGE_KEY = 'esc_countdown_end';
  const COUNTDOWN_DURATION_MS = 48 * 60 * 60 * 1000; // 48 hours in milliseconds

  const hoursEl   = document.getElementById('countdown-hours');
  const minutesEl = document.getElementById('countdown-minutes');
  const secondsEl = document.getElementById('countdown-seconds');

  /**
   * Retrieves (or initialises) the countdown end timestamp from localStorage.
   * If no end time exists or the stored time has already passed, a new
   * 48-hour window is created and persisted.
   */
  function getCountdownEnd() {
    let endTime = parseInt(localStorage.getItem(COUNTDOWN_STORAGE_KEY), 10);

    // If there's no stored end time or it has already passed, reset it
    if (!endTime || isNaN(endTime) || Date.now() >= endTime) {
      endTime = Date.now() + COUNTDOWN_DURATION_MS;
      localStorage.setItem(COUNTDOWN_STORAGE_KEY, endTime.toString());
    }

    return endTime;
  }

  /**
   * Pads a number to two digits (e.g. 5 → "05").
   */
  function pad(n) {
    return String(n).padStart(2, '0');
  }

  /**
   * Calculates the remaining time and updates the DOM elements.
   * Returns false when the countdown has reached zero so the
   * interval can be cleared.
   */
  function updateCountdown() {
    const endTime   = getCountdownEnd();
    const remaining = Math.max(0, endTime - Date.now());

    const totalSeconds = Math.floor(remaining / 1000);
    const hours   = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    if (hoursEl)   hoursEl.textContent   = pad(hours);
    if (minutesEl) minutesEl.textContent = pad(minutes);
    if (secondsEl) secondsEl.textContent = pad(seconds);

    // Return whether the countdown is still running
    return remaining > 0;
  }

  // Only start the timer if the countdown elements exist on the page
  if (hoursEl && minutesEl && secondsEl) {
    // Run immediately so there's no 1-second blank flash
    updateCountdown();

    // Update every second
    const countdownInterval = setInterval(() => {
      const stillRunning = updateCountdown();
      if (!stillRunning) {
        clearInterval(countdownInterval);
      }
    }, 1000);
  }


});
