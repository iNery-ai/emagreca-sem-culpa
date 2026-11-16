/**
 * =========================================================
 * EMAGREÇA SEM CULPA — COMECE PELA ROTINA
 * Script Oficial: Navegação, Modal, FAQ, Timer e Gestão da Oferta
 * =========================================================
 */

document.addEventListener('DOMContentLoaded', () => {

  // =======================================================
  // 1. CONFIGURAÇÃO DA OFERTA E DO CRONÔMETRO
  // =======================================================
  const OFFER_CONFIG = {
    /**
     * Modo de controle do prazo da condição:
     * - 'individual': 90 minutos contados a partir da 1ª visita do usuário (persistido em storage).
     * - 'campaign': Data e hora de término fixas para todos os visitantes da campanha.
     */
    mode: 'individual', // 'individual' | 'campaign'

    /** Duração em minutos no modo individual: 90 min (1h30) */
    durationMinutes: 90,

    /** Data/hora limite no modo 'campaign' (exemplo: '2026-11-20T23:59:59-03:00') */
    campaignDeadline: null,

    /** Preços */
    promoPrice: {
      full: '19,90',
      num: '19',
      cents: ',90'
    },
    regularPrice: {
      full: '47,00',
      num: '47',
      cents: ',00'
    },

    /** Links de Checkout Hotmart */
    checkoutUrlPromo: 'https://pay.hotmart.com/Y106402714O?checkoutMode=10',
    // PENDÊNCIA HOTMART: URL do checkout configurada para R$ 47,00 quando o lote promocional encerrar.
    checkoutUrlRegular: 'https://pay.hotmart.com/Y106402714O?checkoutMode=10',

    storageKeyDeadline: 'esc_offer_deadline_v2',
    storageKeyStart: 'esc_offer_start_v2'
  };


  // =======================================================
  // 2. GESTÃO DO CRONÔMETRO E DOS ESTADOS DA OFERTA
  // =======================================================
  const hoursEl   = document.getElementById('countdown-hours');
  const minutesEl = document.getElementById('countdown-minutes');
  const secondsEl = document.getElementById('countdown-seconds');

  const mainPriceValEl   = document.querySelector('.js-main-price-val');
  const mainPriceCentsEl = document.querySelector('.js-main-price-cents');
  const heroCtaEl        = document.querySelector('.js-hero-cta');
  const offerCtaTextEl   = document.querySelector('.js-offer-cta-text');
  const closingCtaEl     = document.querySelector('.js-closing-cta');
  const stickyPriceEl    = document.querySelector('.js-current-price');
  const stickyTitleEl    = document.querySelector('.js-sticky-title');

  let countdownInterval = null;

  /**
   * Obtém ou inicializa o timestamp de término confiável sem reiniciar no reload.
   */
  function getOfferDeadline() {
    // 1. Verificação se foi forçado via URL para testes (?test_expired=1)
    const urlParams = new URLSearchParams(window.location.search);
    if (urlParams.get('test_expired') === '1') {
      return Date.now() - 1000; // Já expirado
    }

    if (OFFER_CONFIG.mode === 'campaign' && OFFER_CONFIG.campaignDeadline) {
      const parsed = Date.parse(OFFER_CONFIG.campaignDeadline);
      if (!isNaN(parsed)) return parsed;
    }

    // Modo individual: recupera do localStorage
    let storedDeadline = parseInt(localStorage.getItem(OFFER_CONFIG.storageKeyDeadline), 10);
    const now = Date.now();

    if (!storedDeadline || isNaN(storedDeadline)) {
      // Primeira visita: define deadline exatamente para agora + durationMinutes
      storedDeadline = now + (OFFER_CONFIG.durationMinutes * 60 * 1000);
      localStorage.setItem(OFFER_CONFIG.storageKeyDeadline, storedDeadline.toString());
      localStorage.setItem(OFFER_CONFIG.storageKeyStart, now.toString());
    }

    return storedDeadline;
  }

  function padZero(num) {
    return String(Math.max(0, num)).padStart(2, '0');
  }

  /**
   * Atualiza os elementos da página para o estado pós-prazo (R$ 47,00)
   */
  function applyExpiredState() {
    document.body.classList.add('offer-expired');

    if (mainPriceValEl) mainPriceValEl.textContent = OFFER_CONFIG.regularPrice.num;
    if (mainPriceCentsEl) mainPriceCentsEl.textContent = OFFER_CONFIG.regularPrice.cents;

    if (heroCtaEl) heroCtaEl.textContent = `QUERO MEU KIT POR R$${OFFER_CONFIG.regularPrice.full}`;
    if (offerCtaTextEl) offerCtaTextEl.textContent = `QUERO GARANTIR MEU KIT POR R$${OFFER_CONFIG.regularPrice.full}`;
    if (closingCtaEl) closingCtaEl.textContent = `QUERO MEU KIT POR R$${OFFER_CONFIG.regularPrice.full}`;

    if (stickyPriceEl) stickyPriceEl.textContent = `R$ ${OFFER_CONFIG.regularPrice.full}`;
    if (stickyTitleEl) {
      stickyTitleEl.innerHTML = `Kit Digital Completo • <strong class="js-current-price">R$ ${OFFER_CONFIG.regularPrice.full}</strong>`;
    }

    // Atualiza links se houver URL separada para o valor regular
    if (OFFER_CONFIG.checkoutUrlRegular !== OFFER_CONFIG.checkoutUrlPromo) {
      document.querySelectorAll('.js-checkout-link').forEach(link => {
        link.href = OFFER_CONFIG.checkoutUrlRegular;
      });
    }
  }

  /**
   * Atualiza os elementos da página para a condição promocional ativa (R$ 19,90)
   */
  function applyActivePromoState() {
    document.body.classList.remove('offer-expired');

    if (mainPriceValEl) mainPriceValEl.textContent = OFFER_CONFIG.promoPrice.num;
    if (mainPriceCentsEl) mainPriceCentsEl.textContent = OFFER_CONFIG.promoPrice.cents;

    if (heroCtaEl) heroCtaEl.textContent = `QUERO MEU KIT POR R$${OFFER_CONFIG.promoPrice.full}`;
    if (offerCtaTextEl) offerCtaTextEl.textContent = 'QUERO GARANTIR MEU KIT';
    if (closingCtaEl) closingCtaEl.textContent = `QUERO MEU KIT POR R$${OFFER_CONFIG.promoPrice.full}`;

    if (stickyPriceEl) stickyPriceEl.textContent = `R$ ${OFFER_CONFIG.promoPrice.full}`;
  }

  /**
   * Loop do cronômetro regressivo
   */
  function tickTimer() {
    const deadline = getOfferDeadline();
    const now = Date.now();
    const remainingMs = deadline - now;

    if (remainingMs <= 0) {
      if (countdownInterval) clearInterval(countdownInterval);
      if (hoursEl) hoursEl.textContent = '00';
      if (minutesEl) minutesEl.textContent = '00';
      if (secondsEl) secondsEl.textContent = '00';
      applyExpiredState();
      return;
    }

    applyActivePromoState();

    const totalSeconds = Math.floor(remainingMs / 1000);
    const hours = Math.floor(totalSeconds / 3600);
    const minutes = Math.floor((totalSeconds % 3600) / 60);
    const seconds = totalSeconds % 60;

    if (hoursEl) hoursEl.textContent = padZero(hours);
    if (minutesEl) minutesEl.textContent = padZero(minutes);
    if (secondsEl) secondsEl.textContent = padZero(seconds);
  }

  // Inicializa o cronômetro imediatamente
  tickTimer();
  countdownInterval = setInterval(tickTimer, 1000);

  // Helpers no escopo global para facilitar validação e testes
  window.testOfferExpired = function() {
    if (countdownInterval) clearInterval(countdownInterval);
    applyExpiredState();
    console.log('✅ Estado expirado ativado para inspeção visual.');
  };

  window.testOfferReset = function() {
    localStorage.removeItem(OFFER_CONFIG.storageKeyDeadline);
    localStorage.removeItem(OFFER_CONFIG.storageKeyStart);
    applyActivePromoState();
    tickTimer();
    countdownInterval = setInterval(tickTimer, 1000);
    console.log('✅ Oferta reinicializada para 90 minutos.');
  };


  // =======================================================
  // 3. NAVEGAÇÃO E MENU MOBILE
  // =======================================================
  const navToggle = document.getElementById('navToggle');
  const siteNav   = document.getElementById('siteNav');

  if (navToggle && siteNav) {
    navToggle.addEventListener('click', () => {
      const isOpen = siteNav.classList.toggle('is-open');
      navToggle.setAttribute('aria-expanded', isOpen ? 'true' : 'false');
    });

    // Fecha ao clicar em qualquer link da navegação
    siteNav.querySelectorAll('.nav-link, .btn-header').forEach(link => {
      link.addEventListener('click', () => {
        siteNav.classList.remove('is-open');
        navToggle.setAttribute('aria-expanded', 'false');
      });
    });

    // Fecha ao clicar fora
    document.addEventListener('click', (e) => {
      if (!siteNav.contains(e.target) && !navToggle.contains(e.target) && siteNav.classList.contains('is-open')) {
        siteNav.classList.remove('is-open');
        navToggle.setAttribute('aria-expanded', 'false');
      }
    });
  }


  // =======================================================
  // 4. STICKY BAR VISÍVEL NO SCROLL
  // =======================================================
  const stickyBar = document.getElementById('stickyBar');
  const heroSection = document.getElementById('hero');

  function handleScroll() {
    if (!stickyBar || !heroSection) return;
    const heroBottom = heroSection.getBoundingClientRect().bottom;
    // Exibe quando passar do final do hero
    if (heroBottom < 100) {
      stickyBar.classList.add('visible');
    } else {
      stickyBar.classList.remove('visible');
    }
  }

  window.addEventListener('scroll', handleScroll, { passive: true });
  handleScroll();


  // =======================================================
  // 5. ACORDEÃO DE DÚVIDAS FREQUENTES (FAQ ACESSÍVEL)
  // =======================================================
  const faqButtons = document.querySelectorAll('.js-faq-btn');

  faqButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const isExpanded = btn.getAttribute('aria-expanded') === 'true';
      const panel = document.getElementById(btn.getAttribute('aria-controls'));
      const parentItem = btn.closest('.faq-item');

      // Fecha outros itens para foco limpo
      faqButtons.forEach(otherBtn => {
        if (otherBtn !== btn) {
          otherBtn.setAttribute('aria-expanded', 'false');
          const otherPanel = document.getElementById(otherBtn.getAttribute('aria-controls'));
          if (otherPanel) otherPanel.style.maxHeight = null;
          const otherItem = otherBtn.closest('.faq-item');
          if (otherItem) otherItem.classList.remove('active');
        }
      });

      // Alterna o item atual
      if (isExpanded) {
        btn.setAttribute('aria-expanded', 'false');
        if (panel) panel.style.maxHeight = null;
        if (parentItem) parentItem.classList.remove('active');
      } else {
        btn.setAttribute('aria-expanded', 'true');
        if (panel) panel.style.maxHeight = panel.scrollHeight + 'px';
        if (parentItem) parentItem.classList.add('active');
      }
    });
  });


  // =======================================================
  // 6. MODAL LIGHTBOX ACESSÍVEL PARA AMOSTRAS
  // =======================================================
  const sampleData = {
    receita: {
      src: 'samples/amostra-receita.svg',
      title: 'Amostra: Ficha de Receita Real (Frango com legumes na panela — Pág. 11)',
      alt: 'Página de receita real do Guia Educativo: Frango com legumes na panela, ingredientes, preparo e substituições.',
      caption: '*Ficha diagramada real do Guia Educativo (página 11). Exemplo de uso prático.'
    },
    roteiro: {
      src: 'samples/amostra-roteiro.svg',
      title: 'Amostra: Trecho Real do Roteiro de 14 Dias (Dias 01 e 02 — Pág. 18)',
      alt: 'Trecho real do roteiro de atividades do Guia Educativo com tarefas práticas de despensa e campos preenchíveis.',
      caption: '*Ficha do roteiro de 14 dias com campos preenchíveis para digitação ou impressão.'
    },
    planejador: {
      src: 'samples/amostra-planejador.svg',
      title: 'Amostra: Ferramenta 02 — Planejador da Semana (Pág. 26)',
      alt: 'Tabela real do planejador semanal com cardápio simplificado, plano para dias corridos e lista por categorias.',
      caption: '*Ferramenta 02 do Guia e Caderno de Atividades para preenchimento no PDF ou impressão.'
    }
  };

  const sampleModal       = document.getElementById('sampleModal');
  const modalCloseBtn     = document.getElementById('modalCloseBtn');
  const modalFooterClose  = document.querySelector('.js-close-modal-footer');
  const modalTitleEl      = document.getElementById('modalTitle');
  const modalImgEl        = document.getElementById('modalImg');
  const modalCaptionEl    = document.getElementById('modalCaption');
  let lastActiveTrigger   = null;

  function openSampleModal(sampleKey, triggerElement) {
    if (!sampleModal || !sampleData[sampleKey]) return;

    lastActiveTrigger = triggerElement;
    const data = sampleData[sampleKey];

    if (modalTitleEl) modalTitleEl.textContent = data.title;
    if (modalImgEl) {
      modalImgEl.src = data.src;
      modalImgEl.alt = data.alt;
    }
    if (modalCaptionEl) modalCaptionEl.textContent = data.caption;

    sampleModal.classList.add('is-open');
    sampleModal.setAttribute('aria-hidden', 'false');
    document.body.style.overflow = 'hidden';

    // Foco no botão de fechar para acessibilidade
    if (modalCloseBtn) modalCloseBtn.focus();
  }

  function closeSampleModal() {
    if (!sampleModal) return;

    sampleModal.classList.remove('is-open');
    sampleModal.setAttribute('aria-hidden', 'true');
    document.body.style.overflow = '';

    // Retorna foco para o gatilho original
    if (lastActiveTrigger && typeof lastActiveTrigger.focus === 'function') {
      lastActiveTrigger.focus();
    }
  }

  // Gatilhos de abertura
  document.querySelectorAll('.js-open-modal').forEach(btn => {
    btn.addEventListener('click', (e) => {
      e.stopPropagation();
      const sampleKey = btn.getAttribute('data-sample');
      if (sampleKey) openSampleModal(sampleKey, btn);
    });

    btn.addEventListener('keydown', (e) => {
      if (e.key === 'Enter' || e.key === ' ') {
        e.preventDefault();
        const sampleKey = btn.getAttribute('data-sample');
        if (sampleKey) openSampleModal(sampleKey, btn);
      }
    });
  });

  // Gatilhos de fechamento
  if (modalCloseBtn) modalCloseBtn.addEventListener('click', closeSampleModal);
  if (modalFooterClose) modalFooterClose.addEventListener('click', closeSampleModal);

  if (sampleModal) {
    sampleModal.addEventListener('click', (e) => {
      if (e.target === sampleModal) {
        closeSampleModal();
      }
    });
  }

  // Tecla ESC para fechar modal
  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape' && sampleModal && sampleModal.classList.contains('is-open')) {
      closeSampleModal();
    }
  });


  // =======================================================
  // 7. RASTREAMENTO SEGURO DE EVENTOS (SEM DUPLICIDADE)
  // =======================================================
  // Garante que o clique no checkout dispare InitiateCheckout (sem jamais disparar Purchase falso)
  const checkoutLinks = document.querySelectorAll('.js-checkout-link');

  checkoutLinks.forEach(link => {
    link.addEventListener('click', () => {
      if (typeof window.fbq === 'function') {
        const isExpired = document.body.classList.contains('offer-expired');
        const currentVal = isExpired ? 47.00 : 19.90;

        window.fbq('track', 'InitiateCheckout', {
          content_name: 'Emagreça sem Culpa — Comece pela Rotina',
          value: currentVal,
          currency: 'BRL'
        });
      }
    });
  });

});
