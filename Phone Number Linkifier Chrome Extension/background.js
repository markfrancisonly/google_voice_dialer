// background.js (Service Worker for Manifest V3)

importScripts('phonePatterns.js');

const GOOGLE_VOICE_CALL_URL = 'https://voice.google.com/u/0/calls?a=nc,';
const POPUP_WIDTH = 520;
const POPUP_HEIGHT = 680;
const MAX_SELECTED_GOOGLE_VOICE_INPUT_LENGTH = 120;
const MENU_CALL_SELECTION = 'callPhone';

function buildGoogleVoiceCallUrl(tel) {
  return GOOGLE_VOICE_CALL_URL + encodeURIComponent(tel);
}

function getChromeWindow(windowId) {
  return new Promise((resolve) => {
    chrome.windows.get(windowId, (currentWindow) => {
      if (chrome.runtime.lastError) {
        resolve(null);
        return;
      }

      resolve(currentWindow);
    });
  });
}

function createChromeWindow(createData) {
  return new Promise((resolve, reject) => {
    chrome.windows.create(createData, (createdWindow) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      resolve(createdWindow);
    });
  });
}

async function openGoogleVoicePopup(tel, tab) {
  const createData = {
    url: buildGoogleVoiceCallUrl(tel),
    type: 'popup',
    width: POPUP_WIDTH,
    height: POPUP_HEIGHT
  };

  if (tab && typeof tab.windowId === 'number') {
    try {
      const currentWindow = await getChromeWindow(tab.windowId);
      if (
        currentWindow &&
        typeof currentWindow.left === 'number' &&
        typeof currentWindow.top === 'number' &&
        typeof currentWindow.width === 'number' &&
        typeof currentWindow.height === 'number'
      ) {
        createData.left = Math.round(currentWindow.left + (currentWindow.width - POPUP_WIDTH) / 2);
        createData.top = Math.round(currentWindow.top + (currentWindow.height - POPUP_HEIGHT) / 2);
      }
    } catch (err) {
      console.warn('Could not center Google Voice popup:', err);
    }
  }

  await createChromeWindow(createData);
}

function extractTelLinkText(linkUrl) {
  if (!linkUrl || !linkUrl.toLowerCase().startsWith('tel:')) return null;

  let phone = linkUrl.slice(4).trim();
  const queryIndex = phone.search(/[?;]/);
  if (queryIndex !== -1) {
    phone = phone.slice(0, queryIndex);
  }

  try {
    phone = decodeURIComponent(phone);
  } catch (err) {
    // Use the raw phone string if the link contains malformed escaping.
  }

  return phone;
}

function normalizeSelectedPhoneText(phoneText) {
  if (typeof phoneText !== 'string') return null;

  const trimmed = phoneText
    .replace(/[\u2010\u2011\u2012\u2013\u2014\u2212]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();

  if (!trimmed || trimmed.length > MAX_SELECTED_GOOGLE_VOICE_INPUT_LENGTH) return null;
  if (!/\d/.test(trimmed)) return null;
  if (/[^0-9A-Za-z+*#().,:\-\s]/.test(trimmed)) return null;
  if ((trimmed.match(/\+/g) || []).length > 1) return null;

  return trimmed;
}

function normalizeTelLinkPhoneText(phoneText) {
  return PhonePatternSettings.sanitizeGoogleVoicePhoneInput(phoneText, { requireFullMatch: true });
}

async function showInvalidPhoneToast(tab, message) {
  if (!tab || typeof tab.id !== 'number') return;

  try {
    await chrome.scripting.executeScript({
      target: { tabId: tab.id },
      func: (toastMessage) => {
        // Remove any existing toasts first (prevent stacking)
        const oldToasts = document.querySelectorAll('.phone-linkifier-toast');
        oldToasts.forEach(t => t.remove());

        // Create new toast
        const toast = document.createElement('div');
        toast.className = 'phone-linkifier-toast';
        toast.textContent = toastMessage;
        toast.style.cssText = `
          position: fixed;
          bottom: 24px;
          right: 24px;
          background-color: rgba(0, 0, 0, 0.88);
          color: white;
          padding: 12px 20px;
          border-radius: 8px;
          font-family: system-ui, -apple-system, sans-serif;
          font-size: 14px;
          z-index: 999999999;
          box-shadow: 0 4px 20px rgba(0,0,0,0.4);
          max-width: 380px;
          line-height: 1.4;
          pointer-events: none;
          transition: opacity 0.4s ease;
        `;
        document.body.appendChild(toast);

        // Fade out and remove after 4 seconds
        setTimeout(() => {
          toast.style.opacity = '0';
          setTimeout(() => toast.remove(), 500);
        }, 4000);
      },
      args: [message]
    });
  } catch (err) {
    console.warn('Could not show invalid phone number toast:', err);
  }
}

function registerContextMenus() {
  chrome.contextMenus.removeAll(() => {
    chrome.contextMenus.create({
      id: MENU_CALL_SELECTION,
      title: "Call '%s'",
      contexts: ["selection", "link"],
      visible: false
    });
  });
}

registerContextMenus();

function updateContextMenuVisibility(context, shouldRefresh) {
  const hasSelection = Boolean(context.selectionText && context.selectionText.trim());
  const hasTelLink = Boolean(context.linkUrl && context.linkUrl.toLowerCase().startsWith('tel:'));

  chrome.contextMenus.update(MENU_CALL_SELECTION, {
    title: hasSelection ? "Call '%s'" : 'Call link phone number',
    visible: hasSelection || hasTelLink
  }, () => {
    void chrome.runtime.lastError;
    if (shouldRefresh && chrome.contextMenus.refresh) {
      chrome.contextMenus.refresh();
    }
  });
}

if (chrome.contextMenus.onShown) {
  chrome.contextMenus.onShown.addListener((info) => {
    updateContextMenuVisibility(info, true);
  });
}

chrome.runtime.onMessage.addListener((message) => {
  if (!message || message.type !== 'phoneLinkifierContextMenu') return;

  updateContextMenuVisibility(message, false);
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId !== MENU_CALL_SELECTION) return;

  if (info.selectionText) {
    const selected = info.selectionText;
    const tel = normalizeSelectedPhoneText(selected);

    if (!tel) {
      await showInvalidPhoneToast(tab, 'Selected text cannot be sent to Google Voice');
      return;
    }

    try {
      await openGoogleVoicePopup(tel, tab);
    } catch (err) {
      console.error('Error opening Google Voice call popup:', err);
    }
    return;
  }

  if (info.linkUrl) {
    const phoneText = extractTelLinkText(info.linkUrl);
    const tel = phoneText ? normalizeTelLinkPhoneText(phoneText) : null;

    if (tel) {
      try {
        await openGoogleVoicePopup(tel, tab);
      } catch (err) {
        console.error('Error opening Google Voice call popup:', err);
      }
    } else {
      await showInvalidPhoneToast(tab, 'This tel link is not a valid phone number');
    }
  }
});
