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


  // ==========================================
  // 4. CHECKOUT MODAL FLOW
  // ==========================================
  const checkoutModal = document.getElementById('checkoutModal');
  const openButtons = document.querySelectorAll('.open-checkout-btn');
  const closeModalBtn = document.getElementById('closeModalBtn');
  const checkoutForm = document.getElementById('checkoutForm');
  const clientEmailInput = document.getElementById('clientEmail');
  const submitCheckoutBtn = document.getElementById('submitCheckoutBtn');
  
  // Checkout step containers
  const stepForm = document.getElementById('stepForm');
  const stepProcessing = document.getElementById('stepProcessing');
  const stepPixQr = document.getElementById('stepPixQr');
  const stepSuccess = document.getElementById('stepSuccess');
  
  // Processing info element
  const processingText = document.getElementById('processingText');
  const successEmailDisplay = document.getElementById('successEmailDisplay');

  // Payment Options selectors
  const payOptButtons = document.querySelectorAll('.pay-opt-btn');
  const pixDetails = document.getElementById('pixDetails');
  const cardDetails = document.getElementById('cardDetails');
  
  let currentPaymentMethod = 'pix'; // default method

  // Show modal
  const openModal = () => {
    checkoutModal.classList.add('active');
    document.body.style.overflow = 'hidden'; // prevent scrolling behind modal
    resetCheckoutSteps();
  };

  // Close modal
  const closeModal = () => {
    checkoutModal.classList.remove('active');
    document.body.style.overflow = ''; // restore scrolling
  };

  // Attach open events to all buy buttons
  openButtons.forEach(btn => btn.addEventListener('click', openModal));

  // Attach close events
  if (closeModalBtn) {
    closeModalBtn.addEventListener('click', closeModal);
  }
  
  // Close if click is outside content card
  checkoutModal.addEventListener('click', (e) => {
    if (e.target === checkoutModal) {
      closeModal();
    }
  });

  // Reset checkout UI steps
  const resetCheckoutSteps = () => {
    stepForm.classList.remove('hidden');
    stepProcessing.classList.add('hidden');
    stepPixQr.classList.add('hidden');
    stepSuccess.classList.add('hidden');
    checkoutForm.reset();
    
    // Reset payment tab to Pix default
    selectPaymentMethod('pix');
  };

  // Payment method selection
  const selectPaymentMethod = (method) => {
    currentPaymentMethod = method;
    
    // Toggle active state on selector buttons
    payOptButtons.forEach(btn => {
      if (btn.getAttribute('data-method') === method) {
        btn.classList.add('active');
      } else {
        btn.classList.remove('active');
      }
    });

    // Toggle forms and main button text
    if (method === 'pix') {
      pixDetails.classList.remove('hidden');
      cardDetails.classList.add('hidden');
      submitCheckoutBtn.textContent = 'Gerar Pix de Pagamento';
      
      // Make card inputs non-required when Pix is chosen
      toggleCardInputsRequired(false);
    } else {
      pixDetails.classList.add('hidden');
      cardDetails.classList.remove('hidden');
      submitCheckoutBtn.textContent = 'Confirmar Pagamento e Pagar';
      
      // Make card inputs required
      toggleCardInputsRequired(true);
    }
  };

  const toggleCardInputsRequired = (isRequired) => {
    const cardNum = document.getElementById('cardNumber');
    const cardExp = document.getElementById('cardExpiry');
    const cardCvv = document.getElementById('cardCvv');
    if (cardNum && cardExp && cardCvv) {
      cardNum.required = isRequired;
      cardExp.required = isRequired;
      cardCvv.required = isRequired;
    }
  };

  // Handle payment tab click
  payOptButtons.forEach(btn => {
    btn.addEventListener('click', () => {
      const method = btn.getAttribute('data-method');
      selectPaymentMethod(method);
    });
  });

  // Handle checkout form submit
  checkoutForm.addEventListener('submit', (e) => {
    e.preventDefault();
    
    const email = clientEmailInput.value.trim();
    if (!email) return;

    // Transition to Processing Step
    stepForm.classList.add('hidden');
    stepProcessing.classList.remove('hidden');
    processingText.textContent = 'Verificando dados...';

    // Simulate Payment Gateway Response
    if (currentPaymentMethod === 'pix') {
      setTimeout(() => {
        processingText.textContent = 'Gerando chave Pix PixPix...';
        setTimeout(() => {
          // Go to Pix QR Screen
          stepProcessing.classList.add('hidden');
          stepPixQr.classList.remove('hidden');
          successEmailDisplay.textContent = email;
        }, 1000);
      }, 1000);
    } else {
      setTimeout(() => {
        processingText.textContent = 'Processando cartão com a adquirente...';
        setTimeout(() => {
          processingText.textContent = 'Aprovando pagamento...';
          setTimeout(() => {
            // Direct to Success Screen
            stepProcessing.classList.add('hidden');
            stepSuccess.classList.remove('hidden');
            successEmailDisplay.textContent = email;
          }, 1000);
        }, 1200);
      }, 1000);
    }
  });


  // ==========================================
  // 5. PIX ACTIONS (COPY KEY & SIMULATION)
  // ==========================================
  const copyPixBtn = document.getElementById('copyPixBtn');
  const pixKeyInput = document.getElementById('pixKeyInput');
  const mockPixPaidBtn = document.getElementById('mockPixPaidBtn');

  // Copy Pix Key
  if (copyPixBtn && pixKeyInput) {
    copyPixBtn.addEventListener('click', () => {
      pixKeyInput.select();
      pixKeyInput.setSelectionRange(0, 99999); // For mobile devices
      
      navigator.clipboard.writeText(pixKeyInput.value)
        .then(() => {
          copyPixBtn.textContent = 'Copiado!';
          copyPixBtn.style.backgroundColor = '#2ECC71';
          setTimeout(() => {
            copyPixBtn.textContent = 'Copiar Código';
            copyPixBtn.style.backgroundColor = '';
          }, 2000);
        })
        .catch(err => {
          console.error('Failed to copy text: ', err);
          // Fallback if Clipboard API fails
          alert('Erro ao copiar. Selecione o texto e copie manualmente.');
        });
    });
  }

  // Simulate Pix paid successfully
  if (mockPixPaidBtn) {
    mockPixPaidBtn.addEventListener('click', () => {
      stepPixQr.classList.add('hidden');
      stepProcessing.classList.remove('hidden');
      processingText.textContent = 'Confirmando recebimento do Pix...';
      
      setTimeout(() => {
        stepProcessing.classList.add('hidden');
        stepSuccess.classList.remove('hidden');
      }, 1500);
    });
  }


  // ==========================================
  // 6. EBOOK DOWNLOAD GENERATOR
  // ==========================================
  const downloadBookBtn = document.getElementById('downloadBookBtn');
  
  if (downloadBookBtn) {
    downloadBookBtn.addEventListener('click', () => {
      // Change text for feedback
      downloadBookBtn.textContent = '📥 BAIXANDO...';
      downloadBookBtn.disabled = true;

      // Simulate a real PDF file download using a text Blob representation
      const mockPdfContent = `%PDF-1.4
%
1 0 obj
<< /Title (Emagrecimento sem Culpa)
   /Author (Emagrecimento sem Culpa)
   /Subject (Metodo das Mulheres que Comem Bem e Emagrecem de Verdade)
>>
endobj
2 0 obj
<< /Type /Catalog /Pages 3 0 R >>
endobj
3 0 obj
<< /Type /Pages /Kids [4 0 R] /Count 1 >>
endobj
4 0 obj
<< /Type /Page /Parent 3 0 R /MediaBox [0 0 595.275 841.889] /Contents 5 0 R /Resources << >> >>
endobj
5 0 obj
<< /Length 120 >>
stream
BT
/F1 24 Tf
70 750 Td (Parabens! Este e o E-book Emagrecimento sem Culpa.) Tj
0 -40 Td (Siga o guia pratico para alcancar o corpo dos seus sonhos!) Tj
ET
endstream
endobj
xref
0 6
0000000000 65535 f 
0000000015 00000 n 
0000000160 00000 n 
0000000210 00000 n 
0000000278 00000 n 
0000000392 00000 n 
trailer
<< /Size 6 /Root 2 0 R /Info 1 0 R >>
startxref
560
%%EOF`;
      
      setTimeout(() => {
        try {
          const blob = new Blob([mockPdfContent], { type: 'application/pdf' });
          const url = URL.createObjectURL(blob);
          const a = document.createElement('a');
          a.href = url;
          a.download = 'Emagrecimento_Sem_Culpa.pdf';
          document.body.appendChild(a);
          a.click();
          document.body.removeChild(a);
          URL.revokeObjectURL(url);
          
          downloadBookBtn.textContent = '✅ E-BOOK BAIXADO!';
        } catch (err) {
          console.error('Download failed: ', err);
          downloadBookBtn.textContent = '❌ ERRO AO BAIXAR';
          downloadBookBtn.disabled = false;
        }
      }, 1500);
    });
  }

});
