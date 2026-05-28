// background.js (Service Worker for Manifest V3)

importScripts('phonePatterns.js');

const GOOGLE_VOICE_CALL_URL = 'https://voice.google.com/u/0/calls?a=nc,';
const POPUP_WIDTH = 520;
const POPUP_HEIGHT = 680;
const MENU_CALL_SELECTION = 'callPhone';
const MENU_CALL_TEL_LINK = 'callTelLink';

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

async function normalizePhoneText(phoneText, options) {
  const settings = await PhonePatternSettings.loadSettings();
  return PhonePatternSettings.sanitizeGoogleVoicePhoneInput(phoneText, options) ||
    PhonePatternSettings.normalizePhoneForTel(phoneText, settings, options);
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
      contexts: ["selection"]
    });

    chrome.contextMenus.create({
      id: MENU_CALL_TEL_LINK,
      title: 'Call link phone number',
      contexts: ["link"],
      visible: false
    });
  });
}

registerContextMenus();

chrome.runtime.onMessage.addListener((message) => {
  if (!message || message.type !== 'phoneLinkifierContextMenu') return;

  const hasSelection = Boolean(message.selectionText && message.selectionText.trim());
  const hasTelLink = Boolean(message.linkUrl && message.linkUrl.toLowerCase().startsWith('tel:'));

  chrome.contextMenus.update(MENU_CALL_TEL_LINK, {
    visible: hasTelLink && !hasSelection
  }, () => {
    if (chrome.runtime.lastError) return;
  });
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId === MENU_CALL_SELECTION && info.selectionText) {
    const selected = info.selectionText;
    const tel = await normalizePhoneText(selected);

    if (tel) {
      try {
        await openGoogleVoicePopup(tel, tab);
      } catch (err) {
        console.error('Error opening Google Voice call popup:', err);
      }
    } else {
      await showInvalidPhoneToast(tab, 'Selected text is not a valid phone number');
    }
    return;
  }

  if (info.menuItemId === MENU_CALL_TEL_LINK && info.linkUrl) {
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
