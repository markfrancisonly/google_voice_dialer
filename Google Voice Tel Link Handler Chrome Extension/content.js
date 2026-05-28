// content.js - Intercept tel: links and delegate Google Voice opening to the extension.

(function () {
  'use strict';

  // Prevent duplicate listeners if script is injected multiple times.
  if (window.__gvTelHandlerInstalled) return;
  window.__gvTelHandlerInstalled = true;

  document.addEventListener('click', function (event) {
    const target = event.target && event.target.nodeType === Node.ELEMENT_NODE
      ? event.target
      : event.target && event.target.parentElement;

    if (!target || typeof target.closest !== 'function') return;

    const link = target.closest('a[href^="tel:"]');
    if (!link) return;

    event.preventDefault();
    event.stopPropagation();

    chrome.runtime.sendMessage({
      type: 'openGoogleVoiceTelLink',
      href: link.href
    }, () => {
      // Reading lastError prevents noisy console output if the extension is reloading.
      void chrome.runtime.lastError;
    });
  }, { capture: true });
})();
