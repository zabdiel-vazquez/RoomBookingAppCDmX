/**
 * Secure Configuration Management
 * ================================
 * All sensitive configuration should be stored in Script Properties
 *
 * SETUP INSTRUCTIONS:
 * 1. Go to Apps Script Editor > Project Settings > Script Properties
 * 2. Add the following properties:
 *    - SLACK_BOT_TOKEN: Your Slack bot OAuth token
 *    - SLACK_ADMIN_ID: Admin Slack user ID
 *    - SLACK_DEFAULT_CHANNEL: Default notification channel ID
 *    - WEB_APP_URL: Your deployed web app URL (optional)
 *    - ADMIN_EMAILS: Comma-separated list of admin emails
 *
 * SECURITY NOTES:
 * - Never hardcode tokens, passwords, or API keys
 * - Use Script Properties for sensitive configuration
 * - Consider Google Cloud Secret Manager for production
 */

/**
 * Get secure configuration from Script Properties
 */
function getSecureConfig() {
  var props = PropertiesService.getScriptProperties();

  var config = {
    slack: {
      botToken: props.getProperty('SLACK_BOT_TOKEN'),
      adminId: props.getProperty('SLACK_ADMIN_ID'),
      defaultChannel: props.getProperty('SLACK_DEFAULT_CHANNEL')
    },
    app: {
      webAppUrl: props.getProperty('WEB_APP_URL') || ScriptApp.getService().getUrl(),
      adminEmails: (props.getProperty('ADMIN_EMAILS') || '').split(',').map(function(e) {
        return e.trim();
      }).filter(Boolean)
    },
    rooms: getRoomConfiguration(),
    workHours: {
      start: 6,
      end: 17,
      timezone: Session.getScriptTimeZone()
    }
  };

  // Validate required config
  if (!config.slack.botToken) {
    throw new Error('SLACK_BOT_TOKEN not configured in Script Properties');
  }

  return config;
}

/**
 * Room configuration - move to Google Sheets for dynamic management
 * For now kept here but should be migrated to database
 */
function getRoomConfiguration() {
  return {
    calendars: {
      A:   'c_18889q127018oil7n02qvmoft5iue@resource.calendar.google.com',
      B:   'c_1883ad9vmfh8ghfol6ns8kkf5bgfa@resource.calendar.google.com',
      C:   'c_1886fv8a851baj3iit4idtqn4ag6k@resource.calendar.google.com',
      D:   'c_188bb89qo6f88hn1kkugivigcsfd2@resource.calendar.google.com',
      E:   'c_1885btid5mpogg5gl176a8rscph3m@resource.calendar.google.com',
      PB1: 'c_1883fqhsm6floisrj0r0fcqmkclee@resource.calendar.google.com',
      PB2: 'c_188a5bts7c3pagcdnugjotg5nk6u8@resource.calendar.google.com',
      PB3: 'c_188bnkchjldt2hhmjnvn6f56a2moi@resource.calendar.google.com'
    },
    labels: {
      A:   'Room Ajolote · 4 people',
      B:   'Room B · 4 people',
      C:   'Room Balam · 9 people',
      D:   'Room Calupoh · 10 people',
      E:   'Room Xolo (9th Floor) · 30 people',
      PB1: 'Phone Booth Cenzontle',
      PB2: 'Phone Booth Tecolote',
      PB3: 'Phone Booth Quetzal'
    }
  };
}

/**
 * Get Slack user ID mapping from Script Properties
 * This avoids hardcoding user information in source code
 */
function getSlackUserMapping(email) {
  if (!email) return null;

  var props = PropertiesService.getScriptProperties();
  var mappingKey = 'SLACK_USER_MAP_' + email.toLowerCase();

  return props.getProperty(mappingKey);
}

/**
 * Set Slack user ID mapping in Script Properties
 * Admin function to populate user mappings securely
 */
function setSlackUserMapping(email, slackUserId) {
  if (!email || !slackUserId) {
    throw new Error('Email and Slack User ID are required');
  }

  // Validate email format
  if (!isValidEmail(email)) {
    throw new Error('Invalid email format: ' + email);
  }

  var props = PropertiesService.getScriptProperties();
  var mappingKey = 'SLACK_USER_MAP_' + email.toLowerCase();

  props.setProperty(mappingKey, slackUserId);
  console.log('Set Slack mapping for ' + maskEmail(email));
}

/**
 * Bulk import Slack user mappings
 * Run this once to migrate from hardcoded mappings
 */
function importSlackUserMappings(mappings) {
  if (!mappings || typeof mappings !== 'object') {
    throw new Error('Mappings must be an object with email: slackUserId pairs');
  }

  var count = 0;
  for (var email in mappings) {
    if (mappings.hasOwnProperty(email)) {
      try {
        setSlackUserMapping(email, mappings[email]);
        count++;
      } catch (error) {
        console.error('Failed to set mapping for ' + email + ':', error.message);
      }
    }
  }

  console.log('Imported ' + count + ' Slack user mappings');
  return { imported: count };
}

/**
 * Validate email format (RFC 5322 compliant)
 */
function isValidEmail(email) {
  if (!email || typeof email !== 'string') return false;

  // More strict email validation
  var emailRegex = /^[a-zA-Z0-9.!#$%&'*+\/=?^_`{|}~-]+@[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(?:\.[a-zA-Z0-9](?:[a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$/;

  if (!emailRegex.test(email)) return false;

  // Additional checks
  if (email.length > 254) return false; // RFC max length
  var parts = email.split('@');
  if (parts[0].length > 64) return false; // Local part max length

  return true;
}

/**
 * Mask email for logging (PII protection)
 */
function maskEmail(email) {
  if (!email || typeof email !== 'string') return '[no email]';

  var parts = email.split('@');
  if (parts.length !== 2) return '[invalid email]';

  var local = parts[0];
  var domain = parts[1];

  // Show first 2 chars and last char of local part
  var maskedLocal = local.length <= 3
    ? local.charAt(0) + '***'
    : local.charAt(0) + local.charAt(1) + '***' + local.charAt(local.length - 1);

  return maskedLocal + '@' + domain;
}

/**
 * Mask sensitive data in objects for logging
 */
function sanitizeForLog(obj) {
  if (!obj) return obj;

  var sanitized = {};

  for (var key in obj) {
    if (obj.hasOwnProperty(key)) {
      var value = obj[key];
      var lowerKey = key.toLowerCase();

      // Mask sensitive fields
      if (lowerKey.includes('email') && typeof value === 'string') {
        sanitized[key] = maskEmail(value);
      } else if (lowerKey.includes('token') || lowerKey.includes('password') || lowerKey.includes('secret')) {
        sanitized[key] = '[REDACTED]';
      } else if (lowerKey.includes('title') || lowerKey.includes('summary')) {
        // Truncate meeting titles to avoid leaking confidential info
        sanitized[key] = typeof value === 'string' && value.length > 30
          ? value.substring(0, 30) + '...'
          : value;
      } else if (typeof value === 'object' && value !== null) {
        sanitized[key] = sanitizeForLog(value);
      } else {
        sanitized[key] = value;
      }
    }
  }

  return sanitized;
}

/**
 * Check if user is admin
 */
function isAdmin(email) {
  if (!email) return false;

  var config = getSecureConfig();
  return config.app.adminEmails.includes(email.toLowerCase());
}

/**
 * Validate room key
 */
function isValidRoomKey(roomKey) {
  if (!roomKey || typeof roomKey !== 'string') return false;

  var config = getRoomConfiguration();
  return config.calendars.hasOwnProperty(roomKey);
}

/**
 * Validate calendar ID format
 */
function isValidCalendarId(calendarId) {
  if (!calendarId || typeof calendarId !== 'string') return false;

  // Google Calendar IDs are either emails or special resource IDs
  var resourcePattern = /^c_[a-z0-9]{26}@resource\.calendar\.google\.com$/;
  var emailPattern = /^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$/;

  return resourcePattern.test(calendarId) || emailPattern.test(calendarId);
}

/**
 * Validate ISO date string
 */
function isValidISODate(dateString) {
  if (!dateString || typeof dateString !== 'string') return false;

  // Check format YYYY-MM-DD or YYYY-MM-DDTHH:MM:SS
  var isoPattern = /^\d{4}-\d{2}-\d{2}(T\d{2}:\d{2}:\d{2})?$/;
  if (!isoPattern.test(dateString)) return false;

  var date = new Date(dateString);
  return !isNaN(date.getTime());
}

/**
 * Sanitize string input (prevent injection)
 */
function sanitizeString(input, maxLength) {
  if (!input) return '';
  if (typeof input !== 'string') return '';

  // Remove control characters and trim
  var sanitized = input.replace(/[\x00-\x1F\x7F]/g, '').trim();

  // Apply max length
  if (maxLength && sanitized.length > maxLength) {
    sanitized = sanitized.substring(0, maxLength);
  }

  return sanitized;
}

/**
 * Rate limiting helper using Cache Service
 */
function checkRateLimit(key, maxRequests, windowSeconds) {
  var cache = CacheService.getScriptCache();
  var cacheKey = 'rate_limit_' + key;

  var current = cache.get(cacheKey);
  var count = current ? parseInt(current) : 0;

  if (count >= maxRequests) {
    return { allowed: false, retryAfter: windowSeconds };
  }

  cache.put(cacheKey, String(count + 1), windowSeconds);
  return { allowed: true, remaining: maxRequests - count - 1 };
}
