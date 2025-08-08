// content.js
// Non-blocking, safer phone-linkifier for Chrome extensions

// Run only in top frame
if (window.top !== window.self) {
  // don't run in iframes
  console.debug('phone-linkifier: running in iframe — skipping');
} else {
  (function () {
    'use strict';

    // Loose-ish regex — we'll validate digits later
    const PHONE_REGEX = /(?:\+?\d[\d\s().-]{5,}\d)/g;

    // Tags/containers to ignore
    const SKIP_TAGS = new Set([
      'A', 'SCRIPT', 'STYLE', 'HEAD', 'NOSCRIPT', 'INPUT', 'TEXTAREA', 'SELECT',
      'OPTION', 'BUTTON', 'CANVAS', 'SVG', 'PRE', 'CODE', 'IFRAME', 'OBJECT',
      'TIME', 'DATA', 'METER', 'PROGRESS', 'MATH'
    ]);

    // Soft limit so we don't explode on massive pages
    const MAX_TEXT_NODES_TO_PROCESS = 25000;

    // Cache for visibility results to avoid repeated getComputedStyle calls
    const visibilityCache = new WeakMap();

    // Used to debounce mutation processing
    let mutationTimer = null;
    const mutationRoots = new Set();
    let processedTextNodeCount = 0;
    let stoppedDueToLimit = false;

    const scheduleIdle = (fn) => {
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(fn, { timeout: 500 });
      } else {
        setTimeout(fn, 50);
      }
    };

    // Quick Luhn check — used to avoid turning credit cards into phone links
    function luhnCheck(digits) {
      let sum = 0;
      let shouldDouble = false;
      for (let i = digits.length - 1; i >= 0; i--) {
        let d = +digits[i];
        if (shouldDouble) {
          d *= 2;
          if (d > 9) d -= 9;
        }
        sum += d;
        shouldDouble = !shouldDouble;
      }
      return sum % 10 === 0;
    }

    function isProbablyCreditCard(digits) {
      // Credit cards: 13-19 digits often; we'll treat 13-19 with Luhn pass as CC
      if (digits.length < 13 || digits.length > 19) return false;
      return luhnCheck(digits);
    }

    function isVisibleEnough(el) {
      // Walk up a few ancestors checking things cheaply and caching results
      let depth = 0;
      while (el && el.nodeType === Node.ELEMENT_NODE && depth < 10) {
        if (visibilityCache.has(el)) {
          if (!visibilityCache.get(el)) return false;
        } else {
          // cheap checks first
          if (el.hidden || el.getAttribute && el.getAttribute('aria-hidden') === 'true') {
            visibilityCache.set(el, false);
            return false;
          }
          // Slightly more expensive check
          try {
            const s = getComputedStyle(el);
            if (s.display === 'none' || s.visibility === 'hidden' || parseFloat(s.opacity) === 0) {
              visibilityCache.set(el, false);
              return false;
            }
            visibilityCache.set(el, true);
          } catch (err) {
            // getComputedStyle might throw on some exotic nodes; assume visible
            visibilityCache.set(el, true);
          }
        }
        el = el.parentElement;
        depth++;
      }
      return true;
    }

    // Normalize phone to digits with optional leading plus
    function normalizePhoneForTel(phone) {
      const hadPlus = phone.trim().startsWith('+');
      const digits = phone.replace(/\D/g, '');
      // phone digits 7..15 is reasonable (E.164 max 15)
      if (digits.length < 7 || digits.length > 15) return null;
      if (isProbablyCreditCard(digits)) return null; // avoid CC-like numbers
      return (hadPlus ? '+' : '') + digits;
    }

    function shouldSkipTextNode(textNode) {
      if (!textNode || !textNode.parentNode) return true;
      const parent = textNode.parentNode;
      if (parent.nodeType !== Node.ELEMENT_NODE) return true;

      const tag = parent.tagName;
      if (SKIP_TAGS.has(tag)) return true;

      // skip contenteditable areas and form controls
      if (parent.isContentEditable) return true;

      // skip if inside an existing <a>
      // check up to a few ancestors (cheap)
      let el = parent;
      let depth = 0;
      while (el && depth < 6) {
        if (el.tagName === 'A') return true;
        el = el.parentElement;
        depth++;
      }

      // skip mostly whitespace nodes
      if (!textNode.data || !textNode.data.trim()) return true;

      return false;
    }

    function linkifyTextNode(textNode) {
      if (stoppedDueToLimit) return;
      if (processedTextNodeCount > MAX_TEXT_NODES_TO_PROCESS) {
        stoppedDueToLimit = true;
        console.warn('phone-linkifier: reached processing limit, stopping further scans');
        return;
      }

      const parent = textNode.parentNode;
      if (!parent || parent.nodeType !== Node.ELEMENT_NODE) return;
      if (shouldSkipTextNode(textNode)) return;
      if (!isVisibleEnough(parent)) return;

      const text = textNode.data;
      if (!text || !text.trim()) return;

      // Avoid very short text chunks
      if (text.length < 6) return;

      // Reset regex state (important when reusing global regex)
      PHONE_REGEX.lastIndex = 0;

      let match;
      let lastIndex = 0;
      const frag = document.createDocumentFragment();
      let anyMatch = false;

      while ((match = PHONE_REGEX.exec(text)) !== null) {
        if (stoppedDueToLimit) break;

        const phoneText = match[0];
        const start = match.index;
        const end = start + phoneText.length;

        // append leading text
        if (start > lastIndex) {
          frag.appendChild(document.createTextNode(text.slice(lastIndex, start)));
        }

        const tel = normalizePhoneForTel(phoneText);
        if (tel) {
          // create anchor
          const a = document.createElement('a');
          a.setAttribute('href', `tel:${tel}`);
          a.textContent = phoneText;
          // mark it so future scans don't try to re-linkify
          a.dataset.telLinkifier = '1';
          frag.appendChild(a);
          anyMatch = true;
        } else {
          // not a validated phone -> plain text
          frag.appendChild(document.createTextNode(phoneText));
        }

        lastIndex = end;
      }

      if (anyMatch) {
        // append trailing text
        if (lastIndex < text.length) {
          frag.appendChild(document.createTextNode(text.slice(lastIndex)));
        }
        try {
          parent.replaceChild(frag, textNode);
        } catch (err) {
          // might fail if DOM changed — ignore
          return;
        }
      }

      processedTextNodeCount++;
    }

    // Collect text nodes under a root element (skip inside anchors & skip small text nodes)
    function collectTextNodes(root) {
      const nodes = [];
      try {
        const walker = document.createTreeWalker(root, NodeFilter.SHOW_TEXT, null);
        let node;
        while ((node = walker.nextNode())) {
          if (shouldSkipTextNode(node)) continue;
          nodes.push(node);
          if (nodes.length + processedTextNodeCount > MAX_TEXT_NODES_TO_PROCESS) break;
        }
      } catch (err) {
        // TreeWalker may throw on unusual roots; ignore
      }
      return nodes;
    }

    // Process nodes in non-blocking batches
    function processNodesInBatches(nodes, batchSize = 500) {
      let i = 0;
      function step() {
        // Process a batch
        const end = Math.min(nodes.length, i + batchSize);
        for (; i < end; i++) {
          linkifyTextNode(nodes[i]);
        }
        if (i < nodes.length && !stoppedDueToLimit) {
          // schedule next slice
          scheduleIdle(step);
        }
      }
      scheduleIdle(step);
    }

    function processRoot(root) {
      if (stoppedDueToLimit) return;
      const nodes = collectTextNodes(root);
      if (nodes.length === 0) return;
      processNodesInBatches(nodes, 300);
    }

    function processDocumentInitial() {
      if (!document.body) return;
      // start scanning the document body but do it in idle to avoid blocking
      scheduleIdle(() => processRoot(document.body));
    }

    // Debounced mutation handler: collect roots and process them once quiet
    function scheduleMutationProcessing() {
      if (mutationTimer) clearTimeout(mutationTimer);
      mutationTimer = setTimeout(() => {
        // Copy roots and clear the set
        const roots = Array.from(mutationRoots);
        mutationRoots.clear();
        for (const r of roots) {
          if (r && r.nodeType === Node.ELEMENT_NODE) {
            processRoot(r);
          } else if (r && r.nodeType === Node.TEXT_NODE) {
            // if we have a text node directly, process its parent
            const parent = r.parentElement;
            if (parent) processRoot(parent);
          }
        }
      }, 120); // 120ms debounce window
    }

    // Observe DOM changes: only collect added nodes and changed text
    const observer = new MutationObserver((mutations) => {
      for (const m of mutations) {
        if (stoppedDueToLimit) break;
        if (m.type === 'childList') {
          for (const added of m.addedNodes) {
            if (!added) continue;
            // For text nodes, add the parent
            if (added.nodeType === Node.TEXT_NODE) {
              mutationRoots.add(added.parentElement || added);
            } else if (added.nodeType === Node.ELEMENT_NODE) {
              // avoid scanning huge document root repeatedly — mark the element
              mutationRoots.add(added);
            }
          }
        } else if (m.type === 'characterData') {
          // text changed
          mutationRoots.add(m.target.parentElement || m.target);
        }
      }
      scheduleMutationProcessing();
    });

    // Start observing after body exists (or wait for DOMContentLoaded)
    function startObserver() {
      try {
        observer.observe(document.body, { childList: true, subtree: true, characterData: true });
      } catch (err) {
        // if something goes wrong, bail silently
        console.warn('phone-linkifier: observer failed to start', err);
      }
    }

    if (document.readyState === 'loading') {
      window.addEventListener('DOMContentLoaded', () => {
        processDocumentInitial();
        startObserver();
      }, { once: true });
    } else {
      processDocumentInitial();
      startObserver();
    }

    // Public for debugging
    window.__phoneLinkifier = {
      resetCounters() {
        processedTextNodeCount = 0;
        stoppedDueToLimit = false;
        visibilityCache.clear && visibilityCache.clear();
      }
    };
  })();
}
