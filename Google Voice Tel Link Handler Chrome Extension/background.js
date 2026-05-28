const GOOGLE_VOICE_HOME_URL = 'https://voice.google.com/';
const GOOGLE_VOICE_CALL_URL = 'https://voice.google.com/u/0/calls?a=nc,';
const GOOGLE_VOICE_URL_PATTERN = 'https://voice.google.com/*';
const TARGET_STORAGE_KEY = 'googleVoiceTarget';
const BOUNDS_STORAGE_KEY = 'googleVoicePopupBounds';
const POPUP_WIDTH = 520;
const POPUP_HEIGHT = 680;
const MIN_POPUP_WIDTH = 320;
const MIN_POPUP_HEIGHT = 360;
const MAX_POPUP_WIDTH = 1200;
const MAX_POPUP_HEIGHT = 1200;

function getStorageValue(key, defaultValue) {
  return new Promise((resolve) => {
    chrome.storage.local.get({ [key]: defaultValue }, (result) => {
      if (chrome.runtime.lastError) {
        resolve(defaultValue);
        return;
      }

      resolve(result[key]);
    });
  });
}

function getFromStorage(defaultValue) {
  return getStorageValue(TARGET_STORAGE_KEY, defaultValue);
}

function saveTarget(tab) {
  if (!tab || typeof tab.id !== 'number') return Promise.resolve();

  return new Promise((resolve) => {
    chrome.storage.local.set({
      [TARGET_STORAGE_KEY]: {
        tabId: tab.id,
        windowId: tab.windowId
      }
    }, resolve);
  });
}

function isUsablePopupBounds(bounds) {
  return (
    bounds &&
    typeof bounds.width === 'number' &&
    typeof bounds.height === 'number' &&
    bounds.width > 0 &&
    bounds.height > 0
  );
}

function clamp(number, min, max) {
  return Math.min(Math.max(number, min), max);
}

function normalizePopupSize(bounds) {
  return {
    width: clamp(Math.round(bounds.width), MIN_POPUP_WIDTH, MAX_POPUP_WIDTH),
    height: clamp(Math.round(bounds.height), MIN_POPUP_HEIGHT, MAX_POPUP_HEIGHT)
  };
}

function savePopupBounds(bounds) {
  if (!isUsablePopupBounds(bounds)) return Promise.resolve();
  const popupSize = normalizePopupSize(bounds);

  return new Promise((resolve) => {
    chrome.storage.local.set({
      [BOUNDS_STORAGE_KEY]: popupSize
    }, resolve);
  });
}

function getTab(tabId) {
  return new Promise((resolve) => {
    chrome.tabs.get(tabId, (tab) => {
      if (chrome.runtime.lastError) {
        resolve(null);
        return;
      }

      resolve(tab);
    });
  });
}

function queryGoogleVoiceTabs() {
  return new Promise((resolve) => {
    chrome.tabs.query({ url: GOOGLE_VOICE_URL_PATTERN }, (tabs) => {
      if (chrome.runtime.lastError) {
        resolve([]);
        return;
      }

      resolve(tabs || []);
    });
  });
}

function updateTab(tabId, updateProperties) {
  return new Promise((resolve, reject) => {
    chrome.tabs.update(tabId, updateProperties, (tab) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      resolve(tab);
    });
  });
}

function focusWindow(windowId) {
  if (typeof windowId !== 'number') return Promise.resolve();

  return new Promise((resolve) => {
    chrome.windows.update(windowId, { focused: true }, () => {
      resolve();
    });
  });
}

function getWindow(windowId) {
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

async function getSavedPopupSize() {
  const savedBounds = await getStorageValue(BOUNDS_STORAGE_KEY, null);
  if (isUsablePopupBounds(savedBounds)) {
    return normalizePopupSize(savedBounds);
  }

  return {
    width: POPUP_WIDTH,
    height: POPUP_HEIGHT
  };
}

async function getCurrentGoogleVoicePopupSize() {
  const storedTarget = await getFromStorage(null);
  if (storedTarget && typeof storedTarget.windowId === 'number') {
    const currentWindow = await getWindow(storedTarget.windowId);
    if (
      currentWindow &&
      currentWindow.type === 'popup' &&
      (!currentWindow.state || currentWindow.state === 'normal') &&
      isUsablePopupBounds(currentWindow)
    ) {
      return normalizePopupSize(currentWindow);
    }
  }

  return null;
}

async function getPreferredPopupSize() {
  return await getCurrentGoogleVoicePopupSize() || await getSavedPopupSize();
}

async function getCenteredPopupBounds(openerWindowId) {
  const popupSize = await getPreferredPopupSize();
  const bounds = {
    width: popupSize.width,
    height: popupSize.height
  };

  if (typeof openerWindowId !== 'number') return bounds;

  const openerWindow = await getWindow(openerWindowId);
  if (
    !openerWindow ||
    typeof openerWindow.left !== 'number' ||
    typeof openerWindow.top !== 'number' ||
    typeof openerWindow.width !== 'number' ||
    typeof openerWindow.height !== 'number'
  ) {
    return bounds;
  }

  bounds.left = Math.round(openerWindow.left + (openerWindow.width - popupSize.width) / 2);
  bounds.top = Math.round(openerWindow.top + (openerWindow.height - popupSize.height) / 2);
  return bounds;
}

function createChromePopupWindow(url, bounds) {
  return new Promise((resolve, reject) => {
    const createData = {
      url,
      type: 'popup',
      width: bounds.width,
      height: bounds.height
    };

    if (typeof bounds.left === 'number') createData.left = bounds.left;
    if (typeof bounds.top === 'number') createData.top = bounds.top;

    chrome.windows.create(createData, (createdWindow) => {
      if (chrome.runtime.lastError) {
        reject(new Error(chrome.runtime.lastError.message));
        return;
      }

      resolve(createdWindow && createdWindow.tabs ? createdWindow.tabs[0] : null);
    });
  });
}

async function createPopupWindow(url, openerWindowId) {
  const bounds = await getCenteredPopupBounds(openerWindowId);

  try {
    return await createChromePopupWindow(url, bounds);
  } catch (err) {
    if (bounds.width === POPUP_WIDTH && bounds.height === POPUP_HEIGHT) throw err;

    return await createChromePopupWindow(url, {
      width: POPUP_WIDTH,
      height: POPUP_HEIGHT
    });
  }
}

async function findLastGoogleVoiceTarget() {
  const storedTarget = await getFromStorage(null);
  if (storedTarget && typeof storedTarget.tabId === 'number') {
    const storedTab = await getTab(storedTarget.tabId);
    if (storedTab) return storedTab;
  }

  const tabs = await queryGoogleVoiceTabs();
  return tabs
    .slice()
    .sort((a, b) => (b.lastAccessed || 0) - (a.lastAccessed || 0))[0] || null;
}

async function focusLastGoogleVoiceWindowOrOpenPopup() {
  const existingTab = await findLastGoogleVoiceTarget();

  if (existingTab) {
    const updatedTab = await updateTab(existingTab.id, { active: true });
    await focusWindow(updatedTab.windowId);
    await saveTarget(updatedTab);
    return updatedTab;
  }

  const createdTab = await createPopupWindow(GOOGLE_VOICE_HOME_URL);
  await saveTarget(createdTab);
  return createdTab;
}

async function openNewGoogleVoiceCallPopup(url, openerWindowId) {
  const createdTab = await createPopupWindow(url, openerWindowId);
  await saveTarget(createdTab);
  return createdTab;
}

function buildGoogleVoiceCallUrl(phone) {
  return GOOGLE_VOICE_CALL_URL + encodeURIComponent(phone);
}

function normalizeTelHref(href) {
  if (!href || !href.toLowerCase().startsWith('tel:')) return null;

  let phone = href.slice(4).trim();
  const parameterIndex = phone.search(/[?;]/);
  if (parameterIndex !== -1) {
    phone = phone.slice(0, parameterIndex);
  }

  try {
    phone = decodeURIComponent(phone);
  } catch (err) {
    // Use the raw phone value if the page has malformed escaping.
  }

  const cleanPhone = phone.replace(/[^\d+]/g, '');
  return cleanPhone || null;
}

chrome.action.onClicked.addListener(() => {
  focusLastGoogleVoiceWindowOrOpenPopup()
    .catch((err) => console.error('Could not open Google Voice:', err));
});

chrome.runtime.onMessage.addListener((message, sender) => {
  if (!message || message.type !== 'openGoogleVoiceTelLink') return;

  const cleanPhone = normalizeTelHref(message.href);
  if (!cleanPhone) return;

  openNewGoogleVoiceCallPopup(
    buildGoogleVoiceCallUrl(cleanPhone),
    sender.tab && sender.tab.windowId
  ).catch((err) => console.error('Could not open Google Voice call:', err));
});

if (chrome.windows && chrome.windows.onBoundsChanged) {
  chrome.windows.onBoundsChanged.addListener((currentWindow) => {
    if (
      !currentWindow ||
      typeof currentWindow.id !== 'number' ||
      currentWindow.type !== 'popup' ||
      (currentWindow.state && currentWindow.state !== 'normal') ||
      !isUsablePopupBounds(currentWindow)
    ) {
      return;
    }

    getFromStorage(null)
      .then((storedTarget) => {
        if (storedTarget && storedTarget.windowId === currentWindow.id) {
          return savePopupBounds(currentWindow);
        }
      })
      .catch((err) => console.error('Could not save Google Voice popup size:', err));
  });
}
