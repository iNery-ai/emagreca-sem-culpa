document.addEventListener('DOMContentLoaded', () => {
  
  // ==========================================
  // 1. STICKY BAR VISIBILITY ON SCROLL
  // ==========================================
  const stickyBar = document.getElementById('stickyBar');
  const heroSection = document.querySelector('.hero');

  const handleScroll = () => {
    if (!stickyBar || !heroSection) return;
    
    // Show sticky bar once user scrolls past 350px or past the hero section
    const heroHeight = heroSection.offsetHeight;
    if (window.scrollY > (heroHeight - 100)) {
      stickyBar.classList.add('visible');
    } else {
      stickyBar.classList.remove('visible');
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

      // Close all other FAQ items first
      document.querySelectorAll('.faq-item').forEach(item => {
        item.classList.remove('active');
        item.querySelector('.faq-answer').style.maxHeight = null;
      });

      // Toggle current item
      if (!isActive) {
        faqItem.classList.add('active');
        // Set max-height to the exact content scrollHeight for CSS animation
        faqAnswer.style.maxHeight = faqAnswer.scrollHeight + 'px';
      }
    });
  });


});
