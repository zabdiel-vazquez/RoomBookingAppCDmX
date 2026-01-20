# Security & Deployment Guide

## 🔒 Security Best Practices

### Critical Security Improvements Implemented

This document outlines the security enhancements made to the Room Booking application and provides deployment instructions.

---

## 1. Configuration Management

### Before (INSECURE ❌)
```javascript
// Hardcoded in source code
const SLACK_BOT_TOKEN = 'xoxb-...';
const ADMIN_SLACK_ID = 'U08L2CVG29W';
```

### After (SECURE ✅)
```javascript
// Stored in Script Properties
var config = getSecureConfig();
var token = config.slack.botToken;
```

### Setup Instructions

1. **Open Apps Script Editor**
   - Go to https://script.google.com
   - Open your project

2. **Configure Script Properties**
   - Click on Project Settings (⚙️ icon)
   - Scroll to "Script Properties"
   - Add the following properties:

   | Property Name | Description | Example |
   |--------------|-------------|---------|
   | `SLACK_BOT_TOKEN` | Your Slack bot OAuth token | `xoxb-...` |
   | `SLACK_ADMIN_ID` | Admin Slack user ID | `UXXXXXXXXXX` |
   | `SLACK_DEFAULT_CHANNEL` | Default notification channel | `CXXXXXXXXXX` |
   | `ADMIN_EMAILS` | Comma-separated admin emails | `admin@yourcompany.com,it@yourcompany.com` |

3. **Import Slack User Mappings**
   - Instead of hardcoding, run the import function:

   ```javascript
   function setupSlackMappings() {
     // One-time import of user mappings
     var mappings = {
       'user1@yourcompany.com': 'UXXXXXXXXXX',
       'user2@yourcompany.com': 'UXXXXXXXXXX'
       // Add all users here
     };

     importSlackUserMappings(mappings);
   }
   ```

   - Run `setupSlackMappings()` once
   - Delete the function after running (don't keep mappings in code)

---

## 2. Input Validation

### Implemented Protections

✅ **Email Validation (RFC 5322 compliant)**
- Prevents email injection attacks
- Max length enforcement (254 chars)
- Domain whitelist support

✅ **Title Sanitization**
- Max 255 characters
- Control character removal
- XSS prevention

✅ **Date/Time Validation**
- ISO format enforcement
- Past date rejection
- Duration limits (15 min - 8 hours)

✅ **Room Key Validation**
- Whitelist-based validation
- Calendar ID format check

### Usage Example

```javascript
function bookRoom(payload) {
  var userEmail = getActiveUserEmail();

  // Validate request
  var validation = validateRequest('bookRoom', payload, userEmail);

  if (!validation.valid) {
    throw new Error('Validation failed: ' + validation.errors.join(', '));
  }

  // Use sanitized payload
  var sanitized = validation.sanitized;
  // ... proceed with booking
}
```

---

## 3. Authorization & Access Control

### User Authorization Checks

```javascript
function validateUserAuthorization(userEmail, bookingData) {
  // 1. Email domain check
  var allowedDomains = ['apollo.io'];

  // 2. Booking quota check (5 per day)
  var quota = checkUserBookingQuota(userEmail, date);

  // 3. Authorization result
  return { authorized: true/false, reason: '...' };
}
```

### Admin-Only Functions

```javascript
function isAdmin(email) {
  var config = getSecureConfig();
  return config.app.adminEmails.includes(email.toLowerCase());
}

// Protect admin functions
function deleteAllBookings() {
  if (!isAdmin(getActiveUserEmail())) {
    throw new Error('Unauthorized: Admin access required');
  }
  // ... admin operation
}
```

---

## 4. PII Protection & Logging

### Email Masking

**Before:**
```javascript
console.log('Booking created by john.doe@apollo.io');
// Logs: "Booking created by john.doe@apollo.io"
```

**After:**
```javascript
console.log('Booking created by ' + maskEmail('john.doe@apollo.io'));
// Logs: "Booking created by jo***e@apollo.io"
```

### Sanitized Logging

```javascript
function logBookingEvent(bookingData) {
  var sanitized = sanitizeForLog(bookingData);
  // Automatically masks emails, tokens, and truncates titles
  console.log(JSON.stringify(sanitized));
}
```

**What gets masked:**
- ✅ Email addresses → `jo***e@apollo.io`
- ✅ Tokens/passwords → `[REDACTED]`
- ✅ Meeting titles → Truncated to 30 chars
- ✅ Sensitive fields → Removed or masked

---

## 5. Rate Limiting

### API Call Protection

```javascript
function checkRateLimit(key, maxRequests, windowSeconds) {
  var cache = CacheService.getScriptCache();
  // Returns: { allowed: true/false, retryAfter: seconds }
}

// Usage
var limit = checkRateLimit('user:' + email, 10, 60); // 10 requests per minute
if (!limit.allowed) {
  throw new Error('Rate limit exceeded. Retry after ' + limit.retryAfter + 's');
}
```

### User Booking Quotas

- **5 bookings per day** per user
- Prevents room hoarding
- Cached with automatic expiration at midnight

---

## 6. Deployment Checklist

### Pre-Deployment

- [ ] Remove all hardcoded credentials from code
- [ ] Configure Script Properties (see Section 1)
- [ ] Import Slack user mappings
- [ ] Test with a non-admin user account
- [ ] Verify email masking in logs
- [ ] Test rate limiting behavior

### Deployment Steps

1. **Deploy as Web App**
   ```
   Apps Script Editor → Deploy → New deployment
   - Type: Web app
   - Execute as: Me
   - Who has access: Anyone within apollo.io
   ```

2. **Save Deployment URL**
   ```
   Add to Script Properties:
   WEB_APP_URL = https://script.google.com/a/macros/apollo.io/s/...
   ```

3. **Test Authentication**
   - Access web app URL
   - Verify it requires Google login
   - Confirm only @apollo.io emails can access

4. **Configure Slack App**
   - Update Slack app's Request URL to your deployment URL
   - Test Slack notifications
   - Verify DMs are sent correctly

### Post-Deployment

- [ ] Monitor Apps Script execution logs for errors
- [ ] Check Slack admin DMs for booking confirmations
- [ ] Verify no PII appears in logs
- [ ] Test booking creation/cancellation
- [ ] Monitor Google Calendar API quota usage

---

## 7. Security Monitoring

### What to Monitor

1. **Apps Script Executions**
   - Go to Apps Script → Executions
   - Look for failed requests
   - Check error messages for attack patterns

2. **API Quota Usage**
   - Google Cloud Console → APIs & Services → Quotas
   - Calendar API quota consumption
   - Alert if approaching limits

3. **Slack Message Failures**
   - Admin DM should log all booking activity
   - Check for missing notifications

4. **Unusual Patterns**
   - Same user creating many bookings
   - Bookings at odd hours
   - Failed authentication attempts

---

## 8. Incident Response

### If Credentials are Compromised

1. **Immediately Revoke Tokens**
   ```
   - Slack: https://api.slack.com/apps → Your App → OAuth & Permissions → Revoke
   - Google: https://myaccount.google.com/permissions
   ```

2. **Rotate Script Properties**
   - Generate new Slack bot token
   - Update `SLACK_BOT_TOKEN` in Script Properties

3. **Review Access Logs**
   - Apps Script → Executions (check for suspicious activity)
   - Slack workspace audit logs

4. **Notify Users**
   - Send message to team about potential compromise
   - Ask users to review their calendar for unauthorized bookings

### If Unauthorized Bookings Detected

1. **Identify the source**
   ```javascript
   // Check organizer email in calendar event
   var event = Calendar.Events.get('primary', eventId);
   console.log('Organizer:', event.organizer.email);
   ```

2. **Cancel unauthorized bookings**
   ```javascript
   cancelBooking({ eventId: 'suspicious_event_id', source: 'Admin Review' });
   ```

3. **Tighten access controls**
   - Add user to blocklist
   - Reduce booking quotas
   - Enable manual approval for certain rooms

---

## 9. Compliance & Audit

### Data Retention

- **Booking confirmation state**: Cleaned up after 7 days
- **Cache data**: Auto-expires (6 hours for Slack IDs, daily for quotas)
- **Calendar events**: Retained per Google Calendar settings
- **Execution logs**: Retained by Apps Script for 28 days

### GDPR Considerations

This app processes:
- ✅ User email addresses (work email, legitimate interest)
- ✅ Meeting titles (necessary for booking functionality)
- ✅ Slack user IDs (service integration)

**User rights:**
- Request data deletion: Delete calendar events + clear cache
- Access data: View own bookings via `getTodayBookings()`
- Opt-out: Remove from Slack user mappings

### Audit Trail

To implement full audit logging:

```javascript
function logAuditEvent(action, details) {
  var sheet = SpreadsheetApp.openById('AUDIT_SHEET_ID').getActiveSheet();
  sheet.appendRow([
    new Date(),
    getActiveUserEmail(),
    action,
    JSON.stringify(sanitizeForLog(details))
  ]);
}

// Usage
logAuditEvent('booking.created', { room: 'A', title: 'Team meeting' });
logAuditEvent('booking.cancelled', { eventId: 'abc123' });
```

---

## 10. Migration Guide

### From Old (Insecure) Version

1. **Backup current deployment**
   - Make a copy of the Apps Script project
   - Export room calendars configuration

2. **Update code files**
   - Add `Config.js` to your project
   - Add `Validation.js` to your project
   - Update `App.js` to use validation functions
   - Update `bot.js` to use secure config

3. **Migrate hardcoded data**
   ```javascript
   // Run once to migrate Slack mappings
   function migrateToSecureConfig() {
     var oldMappings = {
       // Copy from old bot.js DEFAULT_SLACK_USER_OVERRIDES
     };
     importSlackUserMappings(oldMappings);
   }
   ```

4. **Test before going live**
   - Deploy as new version
   - Test with test user account
   - Verify notifications work
   - Confirm no PII in logs

5. **Switch to new version**
   - Update Slack app webhook URL
   - Deprecate old deployment
   - Monitor for issues

---

## 11. FAQ

**Q: Where are Slack tokens stored?**
A: In Script Properties (encrypted at rest by Google)

**Q: Can users see other users' personal event details?**
A: No, only event titles for events they're invited to

**Q: What happens if quota is exceeded?**
A: User gets error message: "Daily booking quota exceeded"

**Q: Are logs encrypted?**
A: Apps Script logs are accessible only to project editors. PII is masked.

**Q: Can I use this for multiple offices?**
A: Yes, but you need to make room configuration dynamic (currently hardcoded)

---

## 12. Next Steps for Enhanced Security

### Recommended Improvements

1. **Migrate to Google Cloud Secret Manager**
   - More secure than Script Properties
   - Automatic rotation support
   - Audit logging built-in

2. **Implement OAuth 2.0 for Web App**
   - Replace session-based auth with token-based
   - Support API access from external clients

3. **Add Database for Audit Trail**
   - Use Google Sheets or Firestore
   - Immutable append-only log
   - Compliance reporting

4. **Implement Real-Time Monitoring**
   - Slack alerts for suspicious activity
   - Daily quota usage reports
   - Failed authentication tracking

5. **Add E2E Encryption for Sensitive Data**
   - Encrypt meeting titles at rest
   - Decrypt only when needed for display

---

## Support

For security concerns or questions:
- Contact IT Security: it@yourcompany.com
- Report vulnerabilities: security@yourcompany.com
- General questions: admin@yourcompany.com

**Last Updated:** January 2026
**Version:** 2.0 (Security Hardened)
