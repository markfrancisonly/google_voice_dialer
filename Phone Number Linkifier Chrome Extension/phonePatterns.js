// Shared phone-pattern settings and matching helpers for the extension.
(function (global) {
  'use strict';

  const STORAGE_KEY = 'phonePatternSettings';

  const PATTERN_DEFINITIONS = [
    {
      id: 'nanp',
      label: 'US / Canada',
      example: '(650) 618-1499',
      description: 'North American numbers with spaces, dots, dashes, parentheses, or digits only.',
      modeSetting: 'nanpMode',
      modes: [
        {
          value: 'strict',
          label: 'Strict'
        },
        {
          value: 'recommended',
          label: 'Recommended'
        },
        {
          value: 'promiscuous',
          label: 'Promiscuous'
        }
      ],
      defaultEnabled: true
    },
    {
      id: 'international',
      label: 'International',
      example: '+44 20 7946 0958',
      description: 'Numbers outside US and Canada that start with a plus sign.',
      defaultEnabled: true
    },
    {
      id: 'local',
      label: 'Local',
      example: '555-1212',
      description: 'Seven-digit local numbers.',
      defaultEnabled: false
    },
    {
      id: 'vanity',
      label: 'Vanity',
      example: '1-800-FLOWERS',
      description: 'Toll-free numbers containing letters.',
      defaultEnabled: true
    }
  ];

  const DEFAULT_SETTINGS = PATTERN_DEFINITIONS.reduce((settings, pattern) => {
    settings[pattern.id] = pattern.defaultEnabled !== false;
    if (pattern.modeSetting) {
      settings[pattern.modeSetting] = 'recommended';
    }
    return settings;
  }, {});

  const MAX_GOOGLE_VOICE_INPUT_LENGTH = 120;
  const PHONE_SEPARATOR_CHARS = '\\s.\\-\\u2010\\u2011\\u2012\\u2013\\u2014\\u2212';
  const PHONE_SEPARATOR = `[${PHONE_SEPARATOR_CHARS}]`;
  const PHONE_PROMISCUOUS_SEPARATOR_CHARS = `${PHONE_SEPARATOR_CHARS}/\\u00b7`;
  const PHONE_PROMISCUOUS_SEPARATOR = `[${PHONE_PROMISCUOUS_SEPARATOR_CHARS}]`;
  const PHONE_PROMISCUOUS_SEPARATOR_RUN = `${PHONE_PROMISCUOUS_SEPARATOR}+`;
  const PHONE_DASH = '[\\-\\u2010\\u2011\\u2012\\u2013\\u2014\\u2212]';
  const NANP_MODE_VALUES = new Set(['strict', 'recommended', 'promiscuous']);
  const NANP_STRICT_PREFIX = `(?:\\+?1${PHONE_SEPARATOR})?`;
  const NANP_STRICT_PAREN = `\\(\\s*\\d{3}\\s*\\)\\s+\\d{3}(?:${PHONE_DASH}|\\.|\\s+)\\d{4}`;
  const NANP_STRICT_DASHES = `\\d{3}${PHONE_DASH}\\d{3}${PHONE_DASH}\\d{4}`;
  const NANP_STRICT_DOTS = '\\d{3}\\.\\d{3}\\.\\d{4}';
  const NANP_STRICT_SPACES = '\\d{3}\\s+\\d{3}\\s+\\d{4}';

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

  const NANP_MATCHERS = {
    strict: {
      id: 'nanp',
      regex: new RegExp(`(?:^|[^\\d+])(${NANP_STRICT_PREFIX}(?:${NANP_STRICT_PAREN}|${NANP_STRICT_DASHES}|${NANP_STRICT_DOTS}|${NANP_STRICT_SPACES}))(?!\\d)`, 'g'),
      captureGroup: 1
    },
    recommended: {
      id: 'nanp',
      regex: new RegExp(`(?:^|[^\\d+])((?:\\+?1${PHONE_SEPARATOR}?)?(?:\\(\\s*\\d{3}\\s*\\)|\\d{3})${PHONE_SEPARATOR}?\\d{3}${PHONE_SEPARATOR}?\\d{4})(?!\\d)`, 'g'),
      captureGroup: 1
    },
    promiscuous: {
      id: 'nanp',
      regex: new RegExp(`(?:^|[^\\d+])((?:\\+?${PHONE_PROMISCUOUS_SEPARATOR}*1${PHONE_PROMISCUOUS_SEPARATOR}*)?(?:(?:\\(\\s*\\d{3}\\s*\\)|\\d{3})${PHONE_PROMISCUOUS_SEPARATOR}*\\d{3}${PHONE_PROMISCUOUS_SEPARATOR}*\\d{4}|\\d{3}${PHONE_PROMISCUOUS_SEPARATOR_RUN}\\d{3}${PHONE_PROMISCUOUS_SEPARATOR_RUN}\\d{4}))(?!\\d)`, 'g'),
      captureGroup: 1
    }
  };

  const MATCHERS = [
    {
      id: 'international',
      regex: new RegExp(`\\+(?!1(?:\\D|$))\\d{1,3}${PHONE_SEPARATOR}(?:\\d{1,4}${PHONE_SEPARATOR}){1,4}\\d{2,4}`, 'g')
    },
    {
      id: 'local',
      regex: new RegExp(`(?:^|[^\\d${PHONE_SEPARATOR_CHARS}])(\\d{3}${PHONE_SEPARATOR}\\d{4})(?![\\d${PHONE_SEPARATOR_CHARS}])`, 'g'),
      captureGroup: 1
    },
    {
      id: 'vanity',
      regex: new RegExp(`(?:\\+?1${PHONE_SEPARATOR}?)?(?:800|888|877|866|855|844|833|822)${PHONE_SEPARATOR}?[A-Za-z0-9]{3}${PHONE_SEPARATOR}?[A-Za-z0-9]{4}`, 'g')
    }
  ];

  function hasPromiscuousNanpSignal(text) {
    return /\+\s+1/.test(text) ||
      /[\/\u00b7]/.test(text) ||
      /\d\s+[.\-\u2010\u2011\u2012\u2013\u2014\u2212]\s+\d/.test(text);
  }

  function getNanpMatcher(settings, text) {
    if (!settings.nanp) return null;
    if (settings.nanpMode !== 'promiscuous') {
      return NANP_MATCHERS[settings.nanpMode] || NANP_MATCHERS.recommended;
    }

    return hasPromiscuousNanpSignal(text)
      ? NANP_MATCHERS.promiscuous
      : NANP_MATCHERS.recommended;
  }

  function mergeSettings(settings) {
    const source = settings || {};
    const merged = Object.assign({}, DEFAULT_SETTINGS, source);

    if (typeof source.nanp !== 'boolean') {
      const legacyNanpKeys = [
        'nanpParentheses',
        'nanpDashes',
        'nanpDots',
        'nanpSpaces',
        'digitOnly'
      ];
      const legacyNanpValues = legacyNanpKeys
        .filter((key) => typeof source[key] === 'boolean')
        .map((key) => source[key]);

      if (legacyNanpValues.length > 0) {
        merged.nanp = legacyNanpValues.some(Boolean);
      }
    }

    if (typeof source.local !== 'boolean' && typeof source.localSevenDigit === 'boolean') {
      merged.local = source.localSevenDigit;
    }

    if (merged.nanpMode === 'minimal') {
      merged.nanpMode = 'strict';
    } else if (
      merged.nanpMode === 'standard' ||
      merged.nanpMode === 'balanced' ||
      merged.nanpMode === 'flexible'
    ) {
      merged.nanpMode = 'recommended';
    }

    if (!NANP_MODE_VALUES.has(merged.nanpMode)) {
      merged.nanpMode = DEFAULT_SETTINGS.nanpMode;
    }

    return merged;
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

    if (patternId === 'nanp') {
      if (!hasBalancedParentheses(trimmed)) return null;
      if (digits.length === 11 && digits.startsWith('1')) {
        return (hadPlus ? '+' : '') + digits;
      }
      if (digits.length === 10 && !hadPlus) {
        return digits;
      }
      return null;
    }

    if (patternId === 'international') {
      if (!hadPlus || digits.startsWith('1')) return null;
      return '+' + digits;
    }

    if (patternId === 'local') {
      return digits.length === 7 ? digits : null;
    }

    if (patternId === 'vanity') {
      if (!(digits.length === 10 || (digits.length === 11 && digits.startsWith('1')))) return null;
    }

    return (hadPlus ? '+' : '') + digits;
  }

  function collectMatches(text, settings) {
    const enabled = mergeSettings(settings);
    const nanpMatcher = getNanpMatcher(enabled, text);
    const matchers = nanpMatcher ? [nanpMatcher, ...MATCHERS] : MATCHERS;
    const matches = [];

    for (const matcher of matchers) {
      if (!enabled[matcher.id]) continue;

      matcher.regex.lastIndex = 0;
      let match;
      while ((match = matcher.regex.exec(text)) !== null) {
        const phoneText = matcher.captureGroup ? match[matcher.captureGroup] : match[0];
        const relativeIndex = matcher.captureGroup ? match[0].indexOf(phoneText) : 0;
        const start = match.index + relativeIndex;
        if (matcher.id === 'nanp' && start > 0 && /\+\s*$/.test(text.slice(Math.max(0, start - 4), start))) {
          continue;
        }
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
