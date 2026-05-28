// Shared phone-pattern settings and matching helpers for the extension.
(function (global) {
  'use strict';

  const STORAGE_KEY = 'phonePatternSettings';

  const PATTERN_DEFINITIONS = [
    {
      id: 'nanpParentheses',
      label: 'Parentheses',
      example: '(555) 123-4567',
      description: 'US/Canada numbers with a parenthesized area code.'
    },
    {
      id: 'nanpDashes',
      label: 'Dashes',
      example: '555-123-4567',
      description: 'US/Canada numbers separated by hyphens.'
    },
    {
      id: 'nanpDots',
      label: 'Dots',
      example: '555.123.4567',
      description: 'US/Canada numbers separated by periods.'
    },
    {
      id: 'nanpSpaces',
      label: 'Spaces',
      example: '555 123 4567',
      description: 'US/Canada numbers separated by spaces.'
    },
    {
      id: 'localSevenDigit',
      label: 'Local',
      example: '555-1212',
      description: 'Seven-digit local numbers.'
    },
    {
      id: 'digitOnly',
      label: 'Digits only',
      example: '5551234567',
      description: 'Ten- or eleven-digit numbers with no separators.'
    },
    {
      id: 'international',
      label: 'International',
      example: '+44 20 7946 0958',
      description: 'International numbers starting with a plus sign.'
    },
    {
      id: 'vanity',
      label: 'Vanity',
      example: '1-800-FLOWERS',
      description: 'Toll-free numbers containing letters.'
    }
  ];

  const DEFAULT_SETTINGS = PATTERN_DEFINITIONS.reduce((settings, pattern) => {
    settings[pattern.id] = true;
    return settings;
  }, {});

  const MAX_GOOGLE_VOICE_INPUT_LENGTH = 120;

  const LETTER_MAP = {
    A: '2', B: '2', C: '2',
    D: '3', E: '3', F: '3',
    G: '4', H: '4', I: '4',
    J: '5', K: '5', L: '5',
    M: '6', N: '6', O: '6',
    P: '7', Q: '7', R: '7', S: '7',
    T: '8', U: '8', V: '8',
    W: '9', X: '9', Y: '9', Z: '9'
  };

  const MATCHERS = [
    {
      id: 'international',
      regex: /\+\d{1,3}[\s.-](?:\d{1,4}[\s.-]){1,4}\d{2,4}/g
    },
    {
      id: 'vanity',
      regex: /(?:\+?1[-.\s]?)?(?:800|888|877|866|855|844|833|822)[-.\s]?[A-Za-z0-9]{3}[-.\s]?[A-Za-z0-9]{4}/g
    },
    {
      id: 'nanpParentheses',
      regex: /(?:\+?1[-.\s]?)?\(\s*\d{3}\s*\)\s*\d{3}[-.\s]?\d{4}/g
    },
    {
      id: 'nanpDashes',
      regex: /(?:\+?1[-.\s]?)?\d{3}-\d{3}-\d{4}/g
    },
    {
      id: 'nanpDots',
      regex: /(?:\+?1[-.\s]?)?\d{3}\.\d{3}\.\d{4}/g
    },
    {
      id: 'nanpSpaces',
      regex: /(?:\+?1\s+)?\d{3}\s+\d{3}\s+\d{4}/g
    },
    {
      id: 'digitOnly',
      regex: /(?:^|[^\d+])(\+?1?\d{10})(?!\d)/g,
      captureGroup: 1
    },
    {
      id: 'localSevenDigit',
      regex: /(?:^|[^\d.-])(\d{3}[-.]\d{4})(?![\d.-])/g,
      captureGroup: 1
    }
  ];

  function mergeSettings(settings) {
    return Object.assign({}, DEFAULT_SETTINGS, settings || {});
  }

  function mapVanityLetters(phone) {
    const upper = phone.toUpperCase();
    let mapped = '';
    for (const char of upper) {
      mapped += LETTER_MAP[char] || char;
    }
    return mapped;
  }

  function luhnCheck(digits) {
    let sum = 0;
    let shouldDouble = false;
    for (let i = digits.length - 1; i >= 0; i--) {
      let digit = Number(digits[i]);
      if (shouldDouble) {
        digit *= 2;
        if (digit > 9) digit -= 9;
      }
      sum += digit;
      shouldDouble = !shouldDouble;
    }
    return sum % 10 === 0;
  }

  function isProbablyCreditCard(digits) {
    if (digits.length < 13 || digits.length > 19) return false;
    return luhnCheck(digits);
  }

  function normalizeMatch(phone, patternId) {
    const trimmed = phone.slice(0, 80).trim();
    const hadPlus = trimmed.startsWith('+');
    const mapped = mapVanityLetters(trimmed);
    const digits = mapped.replace(/\D/g, '');

    if (digits.length < 7 || digits.length > 15) return null;
    if (isProbablyCreditCard(digits)) return null;
    if (patternId === 'digitOnly' && !(digits.length === 10 || digits.length === 11)) return null;
    if (patternId === 'localSevenDigit' && digits.length !== 7) return null;
    if (patternId !== 'international' && hadPlus && digits.length < 11) return null;

    return (hadPlus ? '+' : '') + digits;
  }

  function collectMatches(text, settings) {
    const enabled = mergeSettings(settings);
    const matches = [];

    for (const matcher of MATCHERS) {
      if (!enabled[matcher.id]) continue;

      matcher.regex.lastIndex = 0;
      let match;
      while ((match = matcher.regex.exec(text)) !== null) {
        const phoneText = matcher.captureGroup ? match[matcher.captureGroup] : match[0];
        const relativeIndex = matcher.captureGroup ? match[0].indexOf(phoneText) : 0;
        const start = match.index + relativeIndex;
        const tel = normalizeMatch(phoneText, matcher.id);

        if (tel) {
          matches.push({
            patternId: matcher.id,
            text: phoneText,
            tel,
            start,
            end: start + phoneText.length
          });
        }
      }
    }

    return matches
      .sort((a, b) => a.start - b.start || (b.end - b.start) - (a.end - a.start))
      .filter((match, index, sorted) => {
        for (let i = 0; i < index; i++) {
          const previous = sorted[i];
          if (match.start < previous.end && match.end > previous.start) return false;
        }
        return true;
      });
  }

  function normalizePhoneForTel(phone, settings, options) {
    const trimmed = phone.trim();
    const matches = collectMatches(trimmed, settings);
    if (matches.length === 0) return null;

    if (options && options.requireFullMatch) {
      const fullMatch = matches.find((match) => match.start === 0 && match.end === trimmed.length);
      return fullMatch ? fullMatch.tel : null;
    }

    return matches[0].tel;
  }

  function hasBalancedParentheses(text) {
    let depth = 0;
    for (const char of text) {
      if (char === '(') depth++;
      if (char === ')') depth--;
      if (depth < 0) return false;
    }
    return depth === 0;
  }

  function isTollFreeVanityCandidate(text) {
    return /^(?:\+?1[-.\s]?)?(?:800|888|877|866|855|844|833|822)[-.\s]?[A-Z0-9]{3}[-.\s]?[A-Z0-9]{4}$/i.test(text);
  }

  function passesGoogleVoiceInputPrecheck(phone) {
    if (typeof phone !== 'string') return false;
    const trimmed = phone.trim();
    return Boolean(trimmed) && trimmed.length <= MAX_GOOGLE_VOICE_INPUT_LENGTH;
  }

  function passesGoogleVoiceSanitizedPostcheck(tel) {
    if (typeof tel !== 'string') return false;
    if (!/^\+?\d+$/.test(tel)) return false;
    if (tel.indexOf('+') > 0) return false;

    const digits = tel.replace(/\D/g, '');
    return digits.length >= 7 && digits.length <= 15;
  }

  function sanitizeGoogleVoiceCandidate(phone) {
    const trimmed = phone.trim();
    if (!passesGoogleVoiceInputPrecheck(trimmed)) return null;
    if (!/^(?:\+?\d|\+?\s*\(\s*\d|\(\s*\d)/.test(trimmed)) return null;
    if (!hasBalancedParentheses(trimmed)) return null;

    const plusCount = (trimmed.match(/\+/g) || []).length;
    if (plusCount > 1 || (plusCount === 1 && !trimmed.startsWith('+'))) return null;

    if (/[A-Z]/i.test(trimmed)) {
      if (!isTollFreeVanityCandidate(trimmed)) return null;
    } else if (/[^+\d\s().-]/.test(trimmed)) {
      return null;
    }

    const hadPlus = /^\+/.test(trimmed);
    const mapped = mapVanityLetters(trimmed);
    const digits = mapped.replace(/\D/g, '');
    const shouldUseInternationalPrefix = hadPlus || digits.length > 10;
    const sanitized = (shouldUseInternationalPrefix ? '+' : '') + digits;
    return passesGoogleVoiceSanitizedPostcheck(sanitized) ? sanitized : null;
  }

  function collectGoogleVoiceInputCandidates(text) {
    const candidates = [];
    const numericCandidateRegex = /(?:\+?\d|\+?\s*\(\s*\d|\(\s*\d)[\d\s().-]{5,}\d/g;
    const vanityCandidateRegex = /(?:\+?1[-.\s]?)?(?:800|888|877|866|855|844|833|822)[-.\s]?[A-Za-z0-9]{3}[-.\s]?[A-Za-z0-9]{4}/g;

    for (const regex of [numericCandidateRegex, vanityCandidateRegex]) {
      regex.lastIndex = 0;
      let match;
      while ((match = regex.exec(text)) !== null) {
        candidates.push({
          text: match[0],
          start: match.index
        });
      }
    }

    return candidates
      .sort((a, b) => a.start - b.start || b.text.length - a.text.length)
      .map((candidate) => candidate.text);
  }

  function sanitizeGoogleVoicePhoneInput(phone, options) {
    if (!passesGoogleVoiceInputPrecheck(phone)) return null;
    const trimmed = phone.trim();

    if (options && options.requireFullMatch) {
      return sanitizeGoogleVoiceCandidate(trimmed);
    }

    for (const candidate of collectGoogleVoiceInputCandidates(trimmed)) {
      const tel = sanitizeGoogleVoiceCandidate(candidate);
      if (tel) return tel;
    }

    return sanitizeGoogleVoiceCandidate(trimmed);
  }

  function isValidGoogleVoicePhoneInput(phone, options) {
    return Boolean(sanitizeGoogleVoicePhoneInput(phone, options));
  }

  function normalizePermissivePhoneForTel(phone, options) {
    return sanitizeGoogleVoicePhoneInput(phone, options);
  }

  function loadSettings() {
    return new Promise((resolve) => {
      if (!global.chrome || !chrome.storage || !chrome.storage.sync) {
        resolve(mergeSettings());
        return;
      }

      chrome.storage.sync.get({ [STORAGE_KEY]: DEFAULT_SETTINGS }, (result) => {
        if (chrome.runtime && chrome.runtime.lastError) {
          resolve(mergeSettings());
          return;
        }

        resolve(mergeSettings(result[STORAGE_KEY]));
      });
    });
  }

  function saveSettings(settings) {
    return new Promise((resolve, reject) => {
      if (!global.chrome || !chrome.storage || !chrome.storage.sync) {
        resolve();
        return;
      }

      chrome.storage.sync.set({ [STORAGE_KEY]: mergeSettings(settings) }, () => {
        if (chrome.runtime && chrome.runtime.lastError) {
          reject(new Error(chrome.runtime.lastError.message));
          return;
        }

        resolve();
      });
    });
  }

  global.PhonePatternSettings = {
    STORAGE_KEY,
    PATTERN_DEFINITIONS,
    getDefaultSettings: () => mergeSettings(),
    mergeSettings,
    findPhoneMatches: collectMatches,
    normalizePhoneForTel,
    passesGoogleVoiceInputPrecheck,
    sanitizeGoogleVoicePhoneInput,
    passesGoogleVoiceSanitizedPostcheck,
    isValidGoogleVoicePhoneInput,
    normalizePermissivePhoneForTel,
    loadSettings,
    saveSettings
  };
})(typeof globalThis !== 'undefined' ? globalThis : this);
