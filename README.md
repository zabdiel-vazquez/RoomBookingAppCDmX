# Apollo CDMX Room Booking System

A secure, scalable room booking application for Apollo CDMX offices, featuring:
- 📅 Weekly grid view with drag-and-drop booking
- 🤖 Slack bot integration for notifications and reminders
- 📊 Real-time dashboard for availability
- 🔒 Enterprise-grade security and PII protection
- ⚡ Rate limiting and quota management

---

## 🚀 Quick Start

### Prerequisites

- Google Workspace account with Calendar API access
- Slack workspace with bot token
- Apps Script deployment permissions

### Files Overview

| File | Purpose |
|------|---------|
| **App.js** | Main backend logic for room booking |
| **bot.js** | Slack integration and notifications |
| **UI.html** | Frontend web interface |
| **Config.js** | 🔒 Secure configuration management |
| **Validation.js** | 🔒 Input validation and sanitization |
| **Migration.js** | 🔒 One-time migration to secure config |
| **SECURITY.md** | 🔒 Security guide and best practices |

---

## 🔒 Security First

This application has been **hardened for security** with the following improvements:

### ✅ What's Secure Now

- **No hardcoded credentials** - All sensitive data in Script Properties
- **PII protection** - Email masking in logs
- **Input validation** - Prevents injection attacks
- **Rate limiting** - Protects against abuse
- **Authorization checks** - User quotas and domain whitelisting
- **Audit logging** - Sanitized activity logs

### ⚠️ Before You Deploy

**DO NOT deploy without completing these steps:**

1. **Read [SECURITY.md](./SECURITY.md)** - Complete security guide
2. **Run Migration.js** - Move secrets to Script Properties
3. **Configure Script Properties** - Add your Slack token
4. **Test with non-admin account** - Verify authorization works
5. **Review logs** - Ensure no PII is exposed

---

## 📦 Installation

### Step 1: Set Up Apps Script Project

1. Go to [script.google.com](https://script.google.com)
2. Create a new project: "Room Booking CDMX"
3. Add all `.js` and `.html` files to the project

### Step 2: Enable Required APIs

1. In Apps Script Editor, click on "Services" (+)
2. Add **Google Calendar API** (v3)
3. Save

### Step 3: Run Security Migration

1. Open `Migration.js` in the editor
2. Run `migrateAllSecuritySettings()`
3. Check execution log for success
4. Go to **Project Settings → Script Properties**
5. Update `SLACK_BOT_TOKEN` with your actual token

### Step 4: Configure Room Calendars

Edit `Config.js` → `getRoomConfiguration()`:

```javascript
calendars: {
  A: 'your_room_a_calendar_id@resource.calendar.google.com',
  B: 'your_room_b_calendar_id@resource.calendar.google.com',
  // ...
}
```

### Step 5: Deploy

1. Click **Deploy → New deployment**
2. Select **Web app**
3. Settings:
   - Execute as: **Me**
   - Who has access: **Anyone within apollo.io**
4. Click **Deploy**
5. Copy the deployment URL
6. Add to Script Properties: `WEB_APP_URL = <your_url>`

### Step 6: Configure Slack

1. Go to [api.slack.com/apps](https://api.slack.com/apps)
2. Create or select your Slack app
3. Add bot scopes:
   - `chat:write`
   - `chat:write.public`
   - `im:write`
   - `users:read`
   - `users:read.email`
4. Install app to workspace
5. Copy Bot User OAuth Token
6. Update Script Properties: `SLACK_BOT_TOKEN = xoxb-...`
7. Set Event Subscriptions URL: `<your_deployment_url>?action=slack`

---

## 🔐 Security Configuration

### Script Properties Setup

**Required Properties:**

```
SLACK_BOT_TOKEN          = xoxb-your-slack-bot-token
SLACK_ADMIN_ID           = U08L2CVG29W
SLACK_DEFAULT_CHANNEL    = C09GZJSPDV4
ADMIN_EMAILS             = admin1@apollo.io,admin2@apollo.io
WEB_APP_URL              = https://script.google.com/a/macros/...
```

### Slack User Mappings

Instead of hardcoding, use the migration script:

```javascript
// Run once in Apps Script
function setupMyTeam() {
  var mappings = {
    'user1@apollo.io': 'U12345678',
    'user2@apollo.io': 'U87654321'
  };
  importSlackUserMappings(mappings);
}
```

### Access Control

Edit `Config.js` → `validateUserAuthorization()`:

```javascript
var allowedDomains = ['apollo.io', 'partner-company.com'];
var MAX_BOOKINGS_PER_DAY = 5; // Customize quota
```

---

## 🎯 Features

### 1. Weekly Grid View

- Drag to select time slots
- Resize bookings
- Color-coded availability
- Quick suggestions

### 2. Slack Integration

**Automated Notifications:**
- ⏰ Reminder 15 min before meeting starts
- 🏁 Reminder 10 min before meeting ends
- ☀️ Daily digest at 8:00 AM
- ✅ Booking confirmation DMs

**Commands:**
- `/rooms` - Open booking interface (if configured)

### 3. Smart Suggestions

- Auto-detects personal calendar events without rooms
- Suggests available rooms for upcoming meetings
- One-click room assignment

### 4. Dashboard

- Real-time availability across all rooms
- Today's bookings at a glance
- Auto-refresh every 30 seconds

---

## 🛡️ Security Features

### Input Validation

All user inputs are validated:

```javascript
var validation = validateRequest('bookRoom', payload, userEmail);
if (!validation.valid) {
  throw new Error(validation.errors.join(', '));
}
```

**Protections:**
- ✅ Email format (RFC 5322 compliant)
- ✅ Max title length (255 chars)
- ✅ XSS prevention
- ✅ SQL injection prevention (N/A for this app)
- ✅ Duration limits (15 min - 8 hours)
- ✅ Past date rejection

### Authorization

```javascript
// Domain whitelist
var allowedDomains = ['apollo.io'];

// Quota enforcement
MAX_BOOKINGS_PER_DAY = 5;

// Admin-only functions
if (!isAdmin(userEmail)) {
  throw new Error('Unauthorized');
}
```

### PII Protection

```javascript
// Before
console.log('User: john.doe@apollo.io');

// After
console.log('User: ' + maskEmail('john.doe@apollo.io'));
// Output: "User: jo***e@apollo.io"
```

### Rate Limiting

```javascript
var limit = checkRateLimit('user:' + email, 10, 60);
if (!limit.allowed) {
  throw new Error('Rate limit exceeded');
}
```

---

## 📊 Scalability

### Current Limits

| Resource | Limit | Impact |
|----------|-------|--------|
| Script Properties | 50 KB | Booking state storage |
| Calendar API | 1M queries/day | 10k-100k users supported |
| Slack API | 1 msg/sec | Sequential DM sending |
| Cache | 10 MB | Slack user ID cache |

### Scaling Recommendations

**For 100+ users:**
1. Migrate booking state from Script Properties to **Firestore**
2. Implement batch Calendar API requests
3. Add Redis/Memcache layer for Slack user IDs
4. Use Pub/Sub for async Slack notifications

**For multiple offices:**
1. Make room configuration dynamic (database-backed)
2. Add office selector to UI
3. Implement location-based authorization

**For enterprise:**
1. Migrate to standalone Node.js service
2. Use Google Cloud Secret Manager
3. Implement OAuth 2.0 for API access
4. Add comprehensive audit logging to BigQuery

---

## 🧪 Testing

### Run Tests

```javascript
// In Apps Script Editor

// Test secure config
function testConfig() {
  var config = getSecureConfig();
  console.log('Config loaded:', config.slack.adminId);
}

// Test validation
function testValidation() {
  var result = validateBookingPayload({
    room: 'A',
    title: 'Test Meeting',
    startISO: '2026-01-21T10:00:00',
    endISO: '2026-01-21T11:00:00'
  });
  console.log('Validation:', result);
}

// Test Slack lookup
function testSlackLookup() {
  var slackId = getSlackUserMapping('your.email@apollo.io');
  console.log('Slack ID:', slackId);
}
```

### Manual Testing Checklist

- [ ] Book a room as normal user
- [ ] Book a room as admin
- [ ] Try booking without authentication
- [ ] Test daily quota limit (book 6 times)
- [ ] Cancel a booking
- [ ] Assign room to existing event
- [ ] Check Slack notifications arrive
- [ ] Verify no PII in execution logs
- [ ] Test with invalid inputs (XSS, injection)

---

## 📝 Usage

### Booking a Room

1. Open the web app URL
2. Navigate to desired week
3. Click and drag on a room row to select time
4. Fill in meeting details
5. Click "Book Room"
6. Receive Slack confirmation

### Assigning Room to Existing Event

1. View "Room Suggestions" tab
2. Click "Assign" next to your event
3. Select available room
4. Receive updated calendar invite

### Canceling a Booking

1. Go to "Your Bookings" tab
2. Find your booking
3. Click "Cancel"
4. Confirm cancellation

---

## 🔧 Configuration

### Work Hours

Edit `Config.js`:

```javascript
workHours: {
  start: 6,  // 6:00 AM
  end: 17,   // 5:00 PM
  timezone: Session.getScriptTimeZone()
}
```

### Notification Timing

Edit `bot.js` (search for these values):

```javascript
REMINDER_BEFORE_MINUTES = 15;  // Start reminder
REMINDER_ENDING_MINUTES = 10;  // Ending reminder
DIGEST_HOUR = 8;               // Daily digest time
```

### Slack Channels

Edit Script Properties:

```
SLACK_DEFAULT_CHANNEL = C09GZJSPDV4  // Default notification channel
SLACK_ADMIN_ID = U08L2CVG29W         // Admin user for alerts
```

---

## 🐛 Troubleshooting

### "SLACK_BOT_TOKEN not configured"

**Fix:** Add token to Script Properties
1. Project Settings → Script Properties
2. Add property: `SLACK_BOT_TOKEN = xoxb-...`

### Slack notifications not arriving

**Checklist:**
- [ ] Bot token is valid (test in Slack API console)
- [ ] Bot has `chat:write` and `im:write` scopes
- [ ] User Slack ID is in Script Properties
- [ ] Run `testSlackLookup()` to verify mapping

### "Rate limit exceeded"

**Cause:** Too many API calls
**Fix:** Wait 60 seconds and try again
**Prevention:** Implement caching for week grid queries

### Bookings appear delayed

**Cause:** Calendar API propagation delay
**Fix:** Wait 5-10 seconds and refresh
**Prevention:** Add optimistic UI updates

### Can't book room (always says "conflict")

**Checklist:**
- [ ] Check Calendar API quota (Cloud Console)
- [ ] Verify calendar sharing with service account
- [ ] Test calendar access: `Calendar.Events.list(calendarId, {})`
- [ ] Check for existing events in time slot

---

## 📈 Monitoring

### View Execution Logs

1. Apps Script Editor → Executions
2. Filter by time range
3. Look for errors or warnings
4. Check execution duration (should be < 10s)

### Monitor API Quota

1. [Google Cloud Console](https://console.cloud.google.com)
2. APIs & Services → Quotas
3. Search "Calendar API"
4. Monitor queries per day

### Slack Activity

Admin user receives DMs for all booking activity:
- ✅ Booking created
- ❌ Booking cancelled
- ⚠️ Quota exceeded
- 🤖 Bot errors

---

## 🤝 Contributing

### Code Style

- Use `var` (not `let`/`const`) for Apps Script compatibility
- Prefer `function` declarations over arrow functions
- Add JSDoc comments for all functions
- Run validation on all user inputs

### Security Guidelines

- **NEVER** commit tokens, passwords, or PII
- **ALWAYS** use `maskEmail()` when logging emails
- **VALIDATE** all inputs with `Validation.js` functions
- **SANITIZE** all outputs with `sanitizeForLog()`
- **TEST** with non-admin account

### Pull Request Checklist

- [ ] All hardcoded secrets removed
- [ ] Input validation added
- [ ] Logging sanitized
- [ ] Tests pass
- [ ] SECURITY.md updated (if security-related)
- [ ] No PII in commit messages

---

## 📄 License

Internal use only - Apollo CDMX

---

## 👤 Contact

**Maintainer:** Zabdiel Vazquez
- Email: zabdiel.vazquez@apollo.io
- Slack: @Zabdiel

**Support:**
- IT Team: it@apollo.io
- Security: security@apollo.io

---

## 🗺️ Roadmap

### v2.1 (Q1 2026)
- [ ] Migrate to Firestore for booking state
- [ ] Add OAuth 2.0 authentication
- [ ] Implement API rate limiting middleware
- [ ] Add analytics dashboard

### v2.2 (Q2 2026)
- [ ] Multi-office support
- [ ] Room equipment tracking
- [ ] QR code check-in
- [ ] Meeting room displays integration

### v3.0 (Q3 2026)
- [ ] Migrate to standalone Node.js service
- [ ] Mobile app (React Native)
- [ ] AI-powered room recommendations
- [ ] Integration with MS Teams

---

**Last Updated:** January 2026
**Version:** 2.0.0 (Security Hardened)
