const cryptoProvider = typeof globalThis !== 'undefined' && globalThis.crypto
  ? globalThis.crypto
  : (typeof window !== 'undefined' && window.crypto ? window.crypto : null);

const SENSITIVE_KEYS = new Set([
  'em', 'ph', 'fn', 'ln', 'ge', 'db', 'ct', 'st', 'zp', 'country', 'external_id'
]);

function isSHA256(str) {
  if (typeof str !== 'string') return false;
  return /^[a-f0-9]{64}$/i.test(str.trim());
}

async function sha256(message) {
  if (!message) return '';
  const normalized = message.trim().toLowerCase();

  // Node.js environment optimization
  if (typeof require !== 'undefined') {
    try {
      const crypto = require('crypto');
      return crypto.createHash('sha256').update(normalized).digest('hex');
    } catch (e) {
      // Fallback to Web Crypto API
    }
  }

  // Web Crypto API
  if (cryptoProvider && cryptoProvider.subtle) {
    const msgBuffer = new TextEncoder().encode(normalized);
    const hashBuffer = await cryptoProvider.subtle.digest('SHA-256', msgBuffer);
    const hashArray = Array.from(new Uint8Array(hashBuffer));
    return hashArray.map(b => b.toString(16).padStart(2, '0')).join('');
  }

  throw new Error('No cryptographic provider found for SHA-256 hashing.');
}

async function sanitizeAndProcessUserData(rawUserData) {
  if (!rawUserData || typeof rawUserData !== 'object') {
    return {};
  }

  const processed = {};

  for (const [key, val] of Object.entries(rawUserData)) {
    // 1. Skip null/undefined
    if (val === null || val === undefined) {
      continue;
    }

    // 2. Handle array and scalar
    const isArray = Array.isArray(val);
    let items = isArray ? val : [val];

    // 3. Filter out null/undefined/empty string items
    items = items.filter(item => item !== null && item !== undefined && String(item).trim() !== '');

    // 4. If no valid items, omit the parameter
    if (items.length === 0) {
      continue;
    }

    if (SENSITIVE_KEYS.has(key)) {
      // Sensitive fields must be hashed and returned as an array of strings
      const hashedItems = await Promise.all(
        items.map(async item => {
          const itemStr = String(item).trim();
          if (isSHA256(itemStr)) {
            return itemStr.toLowerCase();
          }
          return await sha256(itemStr);
        })
      );
      processed[key] = hashedItems;
    } else {
      // Non-sensitive fields: preserve the original format (array or scalar)
      const cleanedItems = items.map(item => typeof item === 'string' ? item.trim() : item);
      processed[key] = isArray ? cleanedItems : cleanedItems[0];
    }
  }

  return processed;
}

function processCustomData(customDataInput) {
  const customData = {
    currency: 'BRL',
    ...customDataInput
  };

  if (customData.value !== undefined && customData.value !== null) {
    const parsedValue = parseFloat(customData.value);
    if (!isNaN(parsedValue)) {
      customData.value = parsedValue;
    } else {
      delete customData.value;
    }
  }

  if (customData.currency) {
    customData.currency = String(customData.currency).trim().toUpperCase();
  }

  return customData;
}

async function generateMetaCAPIPayload(eventName, rawUserData, customDataInput, actionSource = 'website') {
  const sanitizedUserData = await sanitizeAndProcessUserData(rawUserData);
  const processedCustomData = processCustomData(customDataInput);
  const eventTime = Math.floor(Date.now() / 1000);

  return {
    data: [
      {
        event_name: eventName || 'Purchase',
        event_time: eventTime,
        action_source: actionSource,
        user_data: sanitizedUserData,
        custom_data: processedCustomData
      }
    ]
  };
}

// Export module for Node and Browser environments
const CAPI = {
  sha256,
  isSHA256,
  sanitizeAndProcessUserData,
  processCustomData,
  generateMetaCAPIPayload
};

if (typeof module !== 'undefined' && module.exports) {
  module.exports = CAPI;
}

if (typeof window !== 'undefined') {
  window.CAPI = CAPI;
}
