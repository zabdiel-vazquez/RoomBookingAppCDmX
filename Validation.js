/**
 * Input Validation Module
 * =======================
 * Centralized validation for all user inputs
 * Prevents injection attacks, data corruption, and invalid operations
 */

/**
 * Validate booking payload
 * @param {Object} payload - Booking request payload
 * @returns {Object} - { valid: boolean, errors: string[], sanitized: Object }
 */
function validateBookingPayload(payload) {
  var errors = [];
  var sanitized = {};

  // Validate room
  if (!payload || !payload.room) {
    errors.push('Room is required');
  } else if (!isValidRoomKey(payload.room)) {
    errors.push('Invalid room: ' + payload.room);
  } else {
    sanitized.room = payload.room;
  }

  // Validate title
  if (!payload.title || !payload.title.trim()) {
    errors.push('Meeting title is required');
  } else {
    var title = sanitizeString(payload.title, 255);
    if (title.length < 3) {
      errors.push('Title must be at least 3 characters');
    } else if (title.length > 255) {
      errors.push('Title must not exceed 255 characters');
    } else {
      sanitized.title = title;
    }
  }

  // Validate start time
  if (!payload.startISO) {
    errors.push('Start time is required');
  } else if (!isValidISODate(payload.startISO)) {
    errors.push('Invalid start time format');
  } else {
    var startDate = new Date(payload.startISO);
    if (isNaN(startDate.getTime())) {
      errors.push('Invalid start time');
    } else if (startDate < new Date()) {
      errors.push('Cannot book in the past');
    } else {
      sanitized.startISO = payload.startISO;
      sanitized.startDate = startDate;
    }
  }

  // Validate end time
  if (!payload.endISO) {
    errors.push('End time is required');
  } else if (!isValidISODate(payload.endISO)) {
    errors.push('Invalid end time format');
  } else {
    var endDate = new Date(payload.endISO);
    if (isNaN(endDate.getTime())) {
      errors.push('Invalid end time');
    } else if (sanitized.startDate && endDate <= sanitized.startDate) {
      errors.push('End time must be after start time');
    } else {
      sanitized.endISO = payload.endISO;
      sanitized.endDate = endDate;
    }
  }

  // Validate duration (max 8 hours per booking)
  if (sanitized.startDate && sanitized.endDate) {
    var durationMs = sanitized.endDate.getTime() - sanitized.startDate.getTime();
    var durationHours = durationMs / (1000 * 60 * 60);

    if (durationHours > 8) {
      errors.push('Booking duration cannot exceed 8 hours');
    } else if (durationMs < 15 * 60 * 1000) {
      errors.push('Booking duration must be at least 15 minutes');
    }
  }

  // Validate guest email (optional)
  if (payload.guestEmail) {
    var guestEmail = sanitizeString(payload.guestEmail, 254);
    if (guestEmail && !isValidEmail(guestEmail)) {
      errors.push('Invalid guest email format');
    } else if (guestEmail) {
      sanitized.guestEmail = guestEmail;
    }
  }

  // Include optional fields
  if (payload.weekStartISO) sanitized.weekStartISO = payload.weekStartISO;
  if (payload.startCol !== undefined) sanitized.startCol = parseInt(payload.startCol);
  if (payload.steps !== undefined) sanitized.steps = parseInt(payload.steps);

  return {
    valid: errors.length === 0,
    errors: errors,
    sanitized: sanitized
  };
}

/**
 * Validate week grid request payload
 */
function validateWeekGridPayload(payload) {
  var errors = [];
  var sanitized = {};

  // Validate weekStartISO
  if (!payload || !payload.weekStartISO) {
    errors.push('weekStartISO is required');
  } else if (!isValidISODate(payload.weekStartISO)) {
    errors.push('Invalid weekStartISO format (use YYYY-MM-DD)');
  } else {
    sanitized.weekStartISO = payload.weekStartISO;
  }

  // Validate optional parameters with defaults
  sanitized.slotMin = payload.slotMin && !isNaN(parseInt(payload.slotMin))
    ? Math.max(15, Math.min(120, parseInt(payload.slotMin)))
    : 30;

  sanitized.workStart = payload.workStart && !isNaN(parseInt(payload.workStart))
    ? Math.max(0, Math.min(23, parseInt(payload.workStart)))
    : 6;

  sanitized.workEnd = payload.workEnd && !isNaN(parseInt(payload.workEnd))
    ? Math.max(1, Math.min(24, parseInt(payload.workEnd)))
    : 17;

  // Validate work hours make sense
  if (sanitized.workEnd <= sanitized.workStart) {
    errors.push('workEnd must be greater than workStart');
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    sanitized: sanitized
  };
}

/**
 * Validate event assignment payload
 */
function validateAssignRoomPayload(payload) {
  var errors = [];
  var sanitized = {};

  // Validate event ID
  if (!payload || !payload.eventId) {
    errors.push('eventId is required');
  } else {
    var eventId = sanitizeString(payload.eventId, 1024);
    if (!eventId) {
      errors.push('Invalid eventId');
    } else {
      sanitized.eventId = eventId;
    }
  }

  // Validate room
  if (!payload.room) {
    errors.push('room is required');
  } else if (!isValidRoomKey(payload.room)) {
    errors.push('Invalid room key');
  } else {
    sanitized.room = payload.room;
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    sanitized: sanitized
  };
}

/**
 * Validate cancellation payload
 */
function validateCancellationPayload(payload) {
  var errors = [];
  var sanitized = {};

  if (!payload || !payload.eventId) {
    errors.push('eventId is required');
  } else {
    var eventId = sanitizeString(payload.eventId, 1024);
    if (!eventId) {
      errors.push('Invalid eventId');
    } else {
      sanitized.eventId = eventId;
    }
  }

  // Optional source field
  if (payload.source) {
    sanitized.source = sanitizeString(payload.source, 50);
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    sanitized: sanitized
  };
}

/**
 * Validate user authorization for booking
 * @param {string} userEmail - Email of user making the request
 * @param {Object} bookingData - Booking details
 * @returns {Object} - { authorized: boolean, reason: string }
 */
function validateUserAuthorization(userEmail, bookingData) {
  if (!userEmail) {
    return { authorized: false, reason: 'User not authenticated' };
  }

  if (!isValidEmail(userEmail)) {
    return { authorized: false, reason: 'Invalid user email' };
  }

  // Check if user's email domain is allowed
  var allowedDomains = ['apollo.io']; // Add your allowed domains
  var userDomain = userEmail.split('@')[1];

  if (!allowedDomains.includes(userDomain)) {
    return { authorized: false, reason: 'Email domain not authorized' };
  }

  // Check booking quota (example: max 5 bookings per day per user)
  if (bookingData && bookingData.startISO) {
    var quota = checkUserBookingQuota(userEmail, bookingData.startISO);
    if (!quota.allowed) {
      return { authorized: false, reason: quota.reason };
    }
  }

  return { authorized: true };
}

/**
 * Check user's booking quota
 * Prevents abuse by limiting bookings per user
 */
function checkUserBookingQuota(userEmail, dateISO) {
  // This is a simplified version. In production, query actual bookings from database
  var cache = CacheService.getScriptCache();
  var cacheKey = 'booking_quota_' + userEmail + '_' + dateISO.substring(0, 10);

  var current = cache.get(cacheKey);
  var count = current ? parseInt(current) : 0;

  var MAX_BOOKINGS_PER_DAY = 5;

  if (count >= MAX_BOOKINGS_PER_DAY) {
    return {
      allowed: false,
      reason: 'Daily booking quota exceeded (max ' + MAX_BOOKINGS_PER_DAY + ' per day)',
      current: count,
      max: MAX_BOOKINGS_PER_DAY
    };
  }

  // Increment counter (expires at end of day)
  var expirationSeconds = getSecondsUntilEndOfDay();
  cache.put(cacheKey, String(count + 1), expirationSeconds);

  return {
    allowed: true,
    current: count + 1,
    max: MAX_BOOKINGS_PER_DAY,
    remaining: MAX_BOOKINGS_PER_DAY - count - 1
  };
}

/**
 * Get seconds until end of day (for cache expiration)
 */
function getSecondsUntilEndOfDay() {
  var now = new Date();
  var endOfDay = new Date(now.getFullYear(), now.getMonth(), now.getDate(), 23, 59, 59);
  return Math.max(1, Math.floor((endOfDay.getTime() - now.getTime()) / 1000));
}

/**
 * Validate and sanitize Slack message content
 */
function validateSlackMessage(channel, text, blocks) {
  var errors = [];
  var sanitized = {};

  // Validate channel
  if (!channel || typeof channel !== 'string') {
    errors.push('Invalid channel ID');
  } else {
    // Slack channel/user IDs start with C, D, or G
    if (!/^[CDG][A-Z0-9]{8,10}$/.test(channel)) {
      errors.push('Invalid Slack channel format');
    } else {
      sanitized.channel = channel;
    }
  }

  // Validate text
  if (text) {
    var sanitizedText = sanitizeString(text, 3000); // Slack limit is 40k but be conservative
    if (sanitizedText.length > 3000) {
      errors.push('Message text too long (max 3000 chars)');
    } else {
      sanitized.text = sanitizedText;
    }
  }

  // Validate blocks (basic validation)
  if (blocks) {
    if (!Array.isArray(blocks)) {
      errors.push('Blocks must be an array');
    } else if (blocks.length > 50) {
      errors.push('Too many blocks (max 50)');
    } else {
      sanitized.blocks = blocks;
    }
  }

  if (!sanitized.text && !sanitized.blocks) {
    errors.push('Either text or blocks must be provided');
  }

  return {
    valid: errors.length === 0,
    errors: errors,
    sanitized: sanitized
  };
}

/**
 * Comprehensive request validation wrapper
 * Use this as the entry point for all API requests
 */
function validateRequest(endpoint, payload, userEmail) {
  var result = { valid: false, errors: [], sanitized: {}, authorization: { authorized: false } };

  // Validate user authentication
  if (!userEmail) {
    result.errors.push('User not authenticated');
    return result;
  }

  // Endpoint-specific validation
  switch (endpoint) {
    case 'bookRoom':
      result = validateBookingPayload(payload);
      if (result.valid) {
        result.authorization = validateUserAuthorization(userEmail, result.sanitized);
        result.valid = result.authorization.authorized;
        if (!result.valid) {
          result.errors.push(result.authorization.reason);
        }
      }
      break;

    case 'getWeekGrid':
      result = validateWeekGridPayload(payload);
      break;

    case 'assignRoomToExistingEvent':
      result = validateAssignRoomPayload(payload);
      if (result.valid) {
        result.authorization = validateUserAuthorization(userEmail, {});
        result.valid = result.authorization.authorized;
        if (!result.valid) {
          result.errors.push(result.authorization.reason);
        }
      }
      break;

    case 'cancelBooking':
      result = validateCancellationPayload(payload);
      break;

    default:
      result.errors.push('Unknown endpoint: ' + endpoint);
  }

  return result;
}
