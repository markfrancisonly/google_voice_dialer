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
    const phonePatterns = window.PhonePatternSettings;
    let activePatternSettings = phonePatterns.getDefaultSettings();

    // Tags/containers to ignore
    const SKIP_TAGS = new Set([
      'A', 'SCRIPT', 'STYLE', 'HEAD', 'NOSCRIPT', 'INPUT', 'TEXTAREA', 'SELECT',
      'OPTION', 'BUTTON', 'CANVAS', 'SVG', 'PRE', 'CODE', 'IFRAME', 'OBJECT',
      'TIME', 'DATA', 'METER', 'PROGRESS', 'MATH'
    ]);

    // Lowered soft limit to reduce memory on massive pages
    const MAX_TEXT_NODES_TO_PROCESS = 15000;

    // Cache for visibility results to avoid repeated getComputedStyle calls
    const visibilityCache = new WeakMap();

    // Used to debounce mutation processing
    let mutationTimer = null;
    const mutationRoots = new Set();
    let scannedTextNodeCount = 0;
    let stoppedDueToLimit = false;

    // Debug flag (set to false for production to remove logs)
    const DEBUG = false;

    const scheduleIdle = (fn) => {
      if (typeof requestIdleCallback === 'function') {
        requestIdleCallback(fn, { timeout: 500 });
      } else {
        setTimeout(fn, 50);
      }
    };

    // Quick Luhn check — used to avoid turning credit cards into phone links
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
      if (scannedTextNodeCount >= MAX_TEXT_NODES_TO_PROCESS) {
        stoppedDueToLimit = true;
        if (DEBUG) console.warn('phone-linkifier: reached processing limit, stopping further scans');
        return;
      }
      scannedTextNodeCount++;

      const parent = textNode.parentNode;
      if (!parent || parent.nodeType !== Node.ELEMENT_NODE) return;
      if (shouldSkipTextNode(textNode)) return;
      if (!isVisibleEnough(parent)) return;

      const text = textNode.data;
      if (!text || !text.trim()) return;

      // Avoid very short text chunks
      if (text.length < 6) return;
      if (!/\d/.test(text)) return;

      const matches = phonePatterns.findPhoneMatches(text, activePatternSettings);
      if (matches.length === 0) return;

      let lastIndex = 0;
      const frag = document.createDocumentFragment();
      let anyMatch = false;

      try {
        for (const match of matches) {
          if (stoppedDueToLimit) break;

          const phoneText = match.text;
          const start = match.start;
          const end = match.end;

          // append leading text
          if (start > lastIndex) {
            frag.appendChild(document.createTextNode(text.slice(lastIndex, start)));
          }

          // create anchor
          const a = document.createElement('a');
          a.setAttribute('href', `tel:${match.tel}`);
          a.textContent = phoneText;
          // mark it so future scans don't try to re-linkify
          a.dataset.telLinkifier = '1';
          frag.appendChild(a);
          anyMatch = true;

          lastIndex = end;
        }
      } catch (err) {
        if (DEBUG) console.warn('phone-linkifier: phone matching error', err);
        return;
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
          if (DEBUG) console.debug('phone-linkifier: replaceChild failed', err);
          return;
        }
      }

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
          if (nodes.length + scannedTextNodeCount >= MAX_TEXT_NODES_TO_PROCESS) break;
        }
      } catch (err) {
        // TreeWalker may throw on unusual roots; ignore
        if (DEBUG) console.debug('phone-linkifier: TreeWalker error', err);
      }
      return nodes;
    }

    // Process nodes in non-blocking batches
    function processNodesInBatches(nodes, batchSize = 250) {
      let i = 0;
      function step(deadline) {
        const hasIdleBudget = deadline && typeof deadline.timeRemaining === 'function';
        const end = Math.min(nodes.length, i + batchSize);

        for (; i < end; i++) {
          if (hasIdleBudget && deadline.timeRemaining() < 4) break;
          linkifyTextNode(nodes[i]);
        }

        if (i < nodes.length && !stoppedDueToLimit) {
          scheduleIdle(step);
        }
      }
      scheduleIdle(step);
    }

    function processRoot(root) {
      if (stoppedDueToLimit) return;
      const nodes = collectTextNodes(root);
      if (nodes.length === 0) return;
      processNodesInBatches(nodes);
    }

    function processDocumentInitial() {
      if (!document.body) return;
      // start scanning the document body but do it in idle to avoid blocking
      scheduleIdle(() => processRoot(document.body));
    }

    function unlinkGeneratedLinks() {
      const links = document.querySelectorAll('a[data-tel-linkifier="1"]');
      for (const link of links) {
        const parent = link.parentNode;
        if (!parent) continue;
        parent.replaceChild(document.createTextNode(link.textContent || ''), link);
        parent.normalize();
      }
    }

    function resetProcessingState() {
      scannedTextNodeCount = 0;
      stoppedDueToLimit = false;
    }

    function getContextMenuLinkUrl(eventTarget) {
      const element = eventTarget && eventTarget.nodeType === Node.ELEMENT_NODE
        ? eventTarget
        : eventTarget && eventTarget.parentElement;

      if (!element || typeof element.closest !== 'function') return '';

      const link = element.closest('a[href^="tel:"]');
      return link ? link.href : '';
    }

    function updateContextMenuForTarget(eventTarget) {
      const selectionText = String(window.getSelection ? window.getSelection() : '').trim();

      chrome.runtime.sendMessage({
        type: 'phoneLinkifierContextMenu',
        selectionText,
        linkUrl: selectionText ? '' : getContextMenuLinkUrl(eventTarget)
      }, () => {
        // Reading lastError prevents noisy console output if the extension is reloading.
        void chrome.runtime.lastError;
      });
    }

    function updateContextMenuForRightClick(event) {
      updateContextMenuForTarget(event.target);
    }

    function updateContextMenuBeforeRightClick(event) {
      if (event.button !== 2) return;
      updateContextMenuForTarget(event.target);
    }

    document.addEventListener('pointerdown', updateContextMenuBeforeRightClick, true);
    document.addEventListener('mousedown', updateContextMenuBeforeRightClick, true);
    document.addEventListener('contextmenu', updateContextMenuForRightClick, true);

    // Debounced mutation handler: collect roots and process them once quiet
    function getProcessableMutationRoot(root) {
      if (!root) return null;
      if (root.nodeType === Node.ELEMENT_NODE) return root;
      if (root.nodeType === Node.TEXT_NODE) return root.parentElement || null;
      return null;
    }

    function collectTopLevelMutationRoots() {
      const uniqueRoots = [];
      const seenRoots = new Set();

      for (const root of mutationRoots) {
        const processableRoot = getProcessableMutationRoot(root);
        if (!processableRoot || seenRoots.has(processableRoot)) continue;

        seenRoots.add(processableRoot);
        uniqueRoots.push(processableRoot);
      }

      return uniqueRoots.filter((root) => {
        for (const other of uniqueRoots) {
          if (root !== other && other.contains(root)) return false;
        }
        return true;
      });
    }

    function scheduleMutationProcessing() {
      if (mutationTimer) clearTimeout(mutationTimer);
      mutationTimer = setTimeout(() => {
        const roots = collectTopLevelMutationRoots();
        mutationRoots.clear();

        for (const r of roots) {
          processRoot(r);
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
            // Skip if this is our own link (prevents potential loops)
            if (added.nodeType === Node.ELEMENT_NODE && added.dataset && added.dataset.telLinkifier) continue;
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
      if (!document.body) {
        if (DEBUG) console.debug('phone-linkifier: document.body not available yet, delaying observer start');
        setTimeout(startObserver, 50); // Retry after a short delay
        return;
      }
      try {
        observer.observe(document.body, { childList: true, subtree: true, characterData: true });
      } catch (err) {
        // if something goes wrong, bail silently
        if (DEBUG) console.warn('phone-linkifier: observer failed to start', err);
      }
    }

    async function initialize() {
      activePatternSettings = await phonePatterns.loadSettings();

      if (document.readyState === 'loading') {
        window.addEventListener('DOMContentLoaded', () => {
          processDocumentInitial();
          startObserver();
        }, { once: true });
      } else {
        processDocumentInitial();
        startObserver();
      }
    }

    if (chrome.storage && chrome.storage.onChanged) {
      chrome.storage.onChanged.addListener((changes, areaName) => {
        if (areaName !== 'sync' || !changes[phonePatterns.STORAGE_KEY]) return;

        activePatternSettings = phonePatterns.mergeSettings(changes[phonePatterns.STORAGE_KEY].newValue);
        unlinkGeneratedLinks();
        resetProcessingState();
        processDocumentInitial();
      });
    }

    initialize().catch((err) => {
      if (DEBUG) console.warn('phone-linkifier: failed to load settings', err);
      activePatternSettings = phonePatterns.getDefaultSettings();
      processDocumentInitial();
      startObserver();
    });

    // Public for debugging
    window.__phoneLinkifier = {
      resetCounters() {
        resetProcessingState();
        visibilityCache.clear && visibilityCache.clear();
      }
    };
  })();
}
