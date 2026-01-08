// background.js (Service Worker for Manifest V3)

// Normalize handles vanity letters by mapping to digits
function normalizePhoneForTel(phone) {
  phone = phone.slice(0, 50).trim(); // Trim to max 50 chars to prevent long selections
  if (phone.length < 7 || !/[-.() ]/.test(phone)) return null; // Enforce min length and delimiter
  const hadPlus = phone.startsWith('+');
  const letterMap = {
    'A': '2', 'B': '2', 'C': '2',
    'D': '3', 'E': '3', 'F': '3',
    'G': '4', 'H': '4', 'I': '4',
    'J': '5', 'K': '5', 'L': '5',
    'M': '6', 'N': '6', 'O': '6',
    'P': '7', 'Q': '7', 'R': '7', 'S': '7',
    'T': '8', 'U': '8', 'V': '8',
    'W': '9', 'X': '9', 'Y': '9', 'Z': '9'
  };
  const upper = phone.toUpperCase();
  let mapped = '';
  for (let c of upper) {
    mapped += letterMap[c] || c;
  }
  const digits = mapped.replace(/\D/g, '');
  if (digits.length < 7 || digits.length > 15) return null;
  return (hadPlus ? '+' : '') + digits;
}

chrome.runtime.onInstalled.addListener(() => {
  chrome.contextMenus.create({
    id: "callPhone",
    title: "Call '%s'",
    contexts: ["selection"]
  });
});

chrome.contextMenus.onClicked.addListener(async (info, tab) => {
  if (info.menuItemId === "callPhone" && info.selectionText) {
    const selected = info.selectionText;
    const tel = normalizePhoneForTel(selected);

    if (tel) {
      // Success: try to dial
      try {
        await chrome.scripting.executeScript({
          target: { tabId: tab.id },
          func: (telUrl) => {
            const a = document.createElement('a');
            a.href = telUrl;
            a.style.display = 'none';
            document.body.appendChild(a);
            a.click();
            document.body.removeChild(a);
          },
          args: [`tel:${tel}`]
        });
      } catch (err) {
        console.error('Error executing tel: script:', err);
      }
    } else {
      // Failure: show in-page toast
      await chrome.scripting.executeScript({
        target: { tabId: tab.id },
        func: () => {
          // Remove any existing toasts first (prevent stacking)
          const oldToasts = document.querySelectorAll('.phone-linkifier-toast');
          oldToasts.forEach(t => t.remove());

          // Create new toast
          const toast = document.createElement('div');
          toast.className = 'phone-linkifier-toast';
          toast.textContent = 'Selected text is not a valid phone number';
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
        }
      });
    }
  }
});