function validateHubSchema() {
  const required = {
    'Driver Index': ['Referral Owner', 'Driver ID', 'Driver Name', 'Email', 'City', 'Stage', 'Next Required Step'],
    'Driver Actions': ['Action ID', 'Driver ID', 'Operational Status', 'Available for Dispatch?'],
    Dispatch: ['Dispatch ID', 'Driver ID', 'Assignment Status'],
    Events: ['Event ID', 'Event Name', 'Start', 'City / Market'],
    Tasks: ['Task ID', 'Owner', 'Task', 'Status'],
    'Email Queue': [
      'Queue ID', 'Recipient Email', 'Subject', 'Body', 'Status', 'Approved?',
      'Referral Owner', 'Sender Email', 'Delivery Provider', 'Provider Message ID'
    ],
    'Outreach Log': ['Outreach ID', 'Timestamp', 'Entity Type', 'Entity ID', 'Outcome'],
    'Fare Sync': ['Sync ID', 'Run At', 'Direction', 'Entity Type', 'Status']
  };
  const hub = getHub_();
  const problems = [];
  Object.keys(required).forEach(sheetName => {
    const sheet = hub.getSheetByName(sheetName);
    if (!sheet) {
      problems.push(`Missing sheet: ${sheetName}`);
      return;
    }
    const headers = mapHeaders_(sheet.getRange(1, 1, 1, Math.max(1, sheet.getLastColumn())).getDisplayValues()[0]);
    required[sheetName].forEach(header => {
      if (headers[header] === undefined) problems.push(`${sheetName}: missing ${header}`);
    });
  });
  const driverSheet = hub.getSheetByName('Driver Index');
  const result = {
    ok: problems.length === 0,
    checkedAt: formatDateTime_(new Date()),
    driverRows: driverSheet ? Math.max(0, driverSheet.getLastRow() - 1) : 0,
    problems
  };
  console.log(JSON.stringify(result));
  return result;
}

function showSchemaValidation() {
  const result = validateHubSchema();
  SpreadsheetApp.getUi().alert(
    result.ok
      ? `Driver portal is ready. ${result.driverRows} driver records are connected.`
      : `Fix these items before deployment:\n\n${result.problems.join('\n')}`
  );
}

function installAutomationTriggers() {
  const handlers = ['generateDriverNextStepQueue', 'syncTasksToCalendar'];
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (handlers.includes(trigger.getHandlerFunction())) ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('generateDriverNextStepQueue').timeBased().atHour(8).everyDays(1).create();
  ScriptApp.newTrigger('syncTasksToCalendar').timeBased().everyMinutes(15).create();
  return {ok: true, handlers};
}

function installReminderSenderTrigger() {
  const status = getReminderProviderStatus_();
  if (!status.ok) throw new Error(status.problems.join(' '));
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendApprovedEmailQueue') ScriptApp.deleteTrigger(trigger);
  });
  ScriptApp.newTrigger('sendApprovedEmailQueue').timeBased().everyMinutes(15).create();
  return {ok: true, installedBy: status.googleAccount, providers: status.providers};
}

function generateDriverNextStepQueue() {
  const hub = getHub_();
  const drivers = hub.getSheetByName('Driver Index');
  const queue = hub.getSheetByName('Email Queue');
  const log = hub.getSheetByName('Outreach Log');
  const driverHeaders = requireHeaders_(drivers, [
    'Referral Owner', 'Driver ID', 'Driver Name', 'Email', 'Reminder Due?', 'Reminder Subject', 'Reminder Message', 'Cadence Days'
  ]);
  const queueHeaders = requireHeaders_(queue, [
    'Queue ID', 'Created At', 'Entity Type', 'Entity ID', 'Recipient Name', 'Recipient Email',
    'Template ID', 'Subject', 'Body', 'Status', 'Approved?', 'Scheduled For', 'Provider Message ID',
    'Sent At', 'Error', 'Created By', 'Referral Owner', 'Sender Email', 'Delivery Provider'
  ]);
  const logHeaders = requireHeaders_(log, ['Timestamp', 'Entity Type', 'Entity ID', 'Campaign', 'Outcome']);
  const driverValues = drivers.getDataRange().getValues();
  const queueValues = queue.getDataRange().getValues();
  const logValues = log.getDataRange().getValues();
  if (driverValues.length < 2) return {ok: true, created: 0};

  const lastSentByDriver = new Map();
  for (let rowIndex = 1; rowIndex < logValues.length; rowIndex += 1) {
    const row = logValues[rowIndex];
    if (String(row[logHeaders['Entity Type']] || '').toLowerCase() !== 'driver') continue;
    if (String(row[logHeaders.Campaign] || '') !== 'DRV-NEXT') continue;
    if (String(row[logHeaders.Outcome] || '').toLowerCase() !== 'sent') continue;
    const driverId = normalizeId_(row[logHeaders['Entity ID']]);
    const sentAt = row[logHeaders.Timestamp];
    if (!driverId || !(sentAt instanceof Date)) continue;
    const previous = lastSentByDriver.get(driverId);
    if (!previous || sentAt.getTime() > previous.getTime()) lastSentByDriver.set(driverId, sentAt);
  }

  const existing = new Set();
  for (let rowIndex = 1; rowIndex < queueValues.length; rowIndex += 1) {
    const row = queueValues[rowIndex];
    const status = String(row[queueHeaders.Status] || '');
    if (status !== 'Draft' && status !== 'Ready') continue;
    existing.add(`${normalizeId_(row[queueHeaders['Entity ID']])}|${cleanText_(row[queueHeaders.Subject], 500)}`);
  }

  const createdBy = Session.getEffectiveUser().getEmail() || 'Apps Script';
  const autoSend = PropertiesService.getScriptProperties().getProperty('DRIVER_REMINDER_AUTO_SEND') === 'YES';
  const rows = [];
  for (let rowIndex = 1; rowIndex < driverValues.length; rowIndex += 1) {
    const row = driverValues[rowIndex];
    const driverId = normalizeId_(row[driverHeaders['Driver ID']]);
    const email = normalizeEmail_(row[driverHeaders.Email]);
    const due = String(row[driverHeaders['Reminder Due?']] || '').toLowerCase() === 'yes';
    const subject = cleanText_(row[driverHeaders['Reminder Subject']], 500);
    const body = normalizeDriverReminderBody_(cleanText_(row[driverHeaders['Reminder Message']], 10000));
    const referralOwner = cleanText_(row[driverHeaders['Referral Owner']], 80);
    const sender = getSenderForOwner_(referralOwner);
    const cadenceDays = Math.max(1, Number(row[driverHeaders['Cadence Days']]) || 7);
    if (!driverId || !email || !due || !subject || !body || !referralOwner || !sender.email || !sender.provider) continue;

    const lastSent = lastSentByDriver.get(driverId);
    if (lastSent && Date.now() - lastSent.getTime() < cadenceDays * 24 * 60 * 60 * 1000) continue;

    const key = `${driverId}|${subject}`;
    if (existing.has(key)) continue;

    const queueRow = new Array(queue.getLastColumn()).fill('');
    queueRow[queueHeaders['Queue ID']] = Utilities.getUuid();
    queueRow[queueHeaders['Created At']] = new Date();
    queueRow[queueHeaders['Entity Type']] = 'Driver';
    queueRow[queueHeaders['Entity ID']] = driverId;
    queueRow[queueHeaders['Recipient Name']] = cleanText_(row[driverHeaders['Driver Name']], 160);
    queueRow[queueHeaders['Recipient Email']] = email;
    queueRow[queueHeaders['Template ID']] = 'DRV-NEXT';
    queueRow[queueHeaders.Subject] = subject;
    queueRow[queueHeaders.Body] = body;
    queueRow[queueHeaders.Status] = autoSend ? 'Ready' : 'Draft';
    queueRow[queueHeaders['Approved?']] = autoSend ? 'Yes' : 'No';
    queueRow[queueHeaders['Created By']] = createdBy;
    queueRow[queueHeaders['Referral Owner']] = referralOwner;
    queueRow[queueHeaders['Sender Email']] = sender.email;
    queueRow[queueHeaders['Delivery Provider']] = sender.provider;
    rows.push(queueRow);
    existing.add(key);
  }

  if (rows.length) queue.getRange(queue.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
  return {ok: true, created: rows.length, autoSend};
}

function getSenderForOwner_(owner) {
  const normalizedOwner = cleanText_(owner, 80).toLowerCase();
  if (normalizedOwner !== 'tag' && normalizedOwner !== 'partner') {
    return {email: '', name: 'Fare Co-op Driver Operations', provider: ''};
  }
  return {
    email: 'chapter-sender@example.org',
    name: 'Fare Co-op Driver Operations',
    provider: 'Brevo'
  };
}

function normalizeDriverReminderBody_(body) {
  return cleanText_(body, 10000)
    .replace(/Fare Co-op Arizona\s*-\s*/gi, 'Fare Co-op — ')
    .replace(/Fare Co-op Arizona/gi, 'Fare Co-op');
}

function getAuthorizedGmailSenderOptions_(senderEmail, senderName) {
  const requested = normalizeEmail_(senderEmail);
  if (!requested) throw new Error('Sender Email is required.');
  const primary = normalizeEmail_(Session.getEffectiveUser().getEmail());
  const aliases = new Set(GmailApp.getAliases().map(normalizeEmail_));
  const options = {name: senderName || 'Fare Co-op Driver Operations', replyTo: requested};
  if (requested === primary) return options;
  if (!aliases.has(requested)) {
    throw new Error(
      `Gmail sender ${requested} is not authorized for ${primary || 'this Apps Script account'}. ` +
      'Run the reminder sender trigger from Partner\'s Gmail account or configure that address as a Gmail send-as identity.'
    );
  }
  options.from = requested;
  return options;
}

function getBrevoApiKey_() {
  const key = String(PropertiesService.getScriptProperties().getProperty('BREVO_API_KEY') || '').trim();
  if (!key) throw new Error('BREVO_API_KEY is missing from Apps Script Project Settings → Script Properties.');
  return key;
}

function validateBrevoConnection_() {
  const response = UrlFetchApp.fetch('https://api.brevo.com/v3/account', {
    method: 'get',
    headers: {'api-key': getBrevoApiKey_(), accept: 'application/json'},
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  if (code < 200 || code >= 300) throw new Error(`Brevo credential check failed with HTTP ${code}.`);
  return {ok: true, httpStatus: code};
}

function sendViaBrevo_(message) {
  const payload = {
    sender: {email: message.senderEmail, name: message.senderName},
    to: [{email: message.recipient, name: message.recipientName || undefined}],
    subject: message.subject,
    textContent: message.body,
    replyTo: {email: message.senderEmail, name: message.senderName}
  };
  const response = UrlFetchApp.fetch('https://api.brevo.com/v3/smtp/email', {
    method: 'post',
    contentType: 'application/json',
    headers: {'api-key': getBrevoApiKey_(), accept: 'application/json'},
    payload: JSON.stringify(payload),
    muteHttpExceptions: true
  });
  const code = response.getResponseCode();
  const text = response.getContentText();
  if (code < 200 || code >= 300) throw new Error(`Brevo send failed with HTTP ${code}: ${cleanText_(text, 500)}`);
  let parsed = {};
  try { parsed = JSON.parse(text || '{}'); } catch (ignored) {}
  return {provider: 'Brevo', messageId: cleanText_(parsed.messageId || parsed.messageIds || '', 500)};
}

function sendViaGmail_(message) {
  const options = getAuthorizedGmailSenderOptions_(message.senderEmail, message.senderName);
  GmailApp.sendEmail(message.recipient, message.subject, message.body, options);
  return {provider: 'Gmail', messageId: ''};
}

function getReminderProviderStatus_() {
  const problems = [];
  const providers = [];
  const googleAccount = normalizeEmail_(Session.getEffectiveUser().getEmail());
  const sender = getSenderForOwner_('Tag');
  try {
    const connection = validateBrevoConnection_();
    providers.push({
      owner: 'Tag + Partner',
      provider: 'Brevo',
      sender: sender.email,
      ready: true,
      httpStatus: connection.httpStatus
    });
  } catch (error) {
    const message = cleanText_(error && error.message ? error.message : error, 500);
    providers.push({owner: 'Tag + Partner', provider: 'Brevo', sender: sender.email, ready: false, error: message});
    problems.push(message);
  }
  return {ok: problems.length === 0, googleAccount, providers, problems};
}

function showReminderProviderStatus() {
  const status = getReminderProviderStatus_();
  const lines = status.providers.map(item =>
    `${item.ready ? 'READY' : 'BLOCKED'} — ${item.owner}: ${item.provider} / ${item.sender}${item.error ? `\n${item.error}` : ''}`
  );
  SpreadsheetApp.getUi().alert(`Reminder sender status\n\n${lines.join('\n\n')}`);
  return status;
}

function activateDriverReminderAutomation() {
  const status = getReminderProviderStatus_();
  if (!status.ok) throw new Error(`Reminder automation was not activated. ${status.problems.join(' ')}`);
  const hub = getHub_();
  const queue = hub.getSheetByName('Email Queue');
  const headers = requireHeaders_(queue, [
    'Template ID', 'Status', 'Approved?', 'Referral Owner', 'Sender Email', 'Delivery Provider'
  ]);
  const values = queue.getDataRange().getValues();
  let approved = 0;
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    if (String(row[headers['Template ID']] || '') !== 'DRV-NEXT') continue;
    if (String(row[headers.Status] || '') !== 'Draft') continue;
    const owner = cleanText_(row[headers['Referral Owner']], 80);
    const sender = getSenderForOwner_(owner);
    if (!sender.email || !sender.provider) continue;
    queue.getRange(rowIndex + 1, headers.Status + 1).setValue('Ready');
    queue.getRange(rowIndex + 1, headers['Approved?'] + 1).setValue('Yes');
    queue.getRange(rowIndex + 1, headers['Sender Email'] + 1).setValue(sender.email);
    queue.getRange(rowIndex + 1, headers['Delivery Provider'] + 1).setValue(sender.provider);
    approved += 1;
  }
  PropertiesService.getScriptProperties().setProperty('DRIVER_REMINDER_AUTO_SEND', 'YES');
  const trigger = installReminderSenderTrigger();
  return {ok: true, approved, trigger};
}

function configureReviewFirstReminderAutomation() {
  const status = getReminderProviderStatus_();
  if (!status.ok) throw new Error(`Review-first reminder automation was not configured. ${status.problems.join(' ')}`);
  PropertiesService.getScriptProperties().setProperty('DRIVER_REMINDER_AUTO_SEND', 'NO');
  const queueResult = generateDriverNextStepQueue();
  const automationTriggers = installAutomationTriggers();
  const senderTrigger = installReminderSenderTrigger();
  return {
    ok: true,
    mode: 'review-first',
    queueResult,
    automationTriggers,
    senderTrigger,
    sender: getSenderForOwner_('Tag')
  };
}

function pauseDriverReminderAutomation() {
  PropertiesService.getScriptProperties().setProperty('DRIVER_REMINDER_AUTO_SEND', 'NO');
  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === 'sendApprovedEmailQueue') ScriptApp.deleteTrigger(trigger);
  });
  return {ok: true, paused: true};
}

function countBrevoSendsToday_(log, headers) {
  const values = log.getDataRange().getValues();
  const today = Utilities.formatDate(new Date(), PORTAL_TIME_ZONE, 'yyyy-MM-dd');
  let count = 0;
  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    const timestamp = row[headers.Timestamp];
    if (!(timestamp instanceof Date)) continue;
    if (Utilities.formatDate(timestamp, PORTAL_TIME_ZONE, 'yyyy-MM-dd') !== today) continue;
    if (String(row[headers.Outcome] || '').toLowerCase() !== 'sent') continue;
    if (!String(row[headers.Notes] || '').includes('Provider: Brevo')) continue;
    count += 1;
  }
  return count;
}

function sendApprovedEmailQueue() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(1000)) return {ok: false, message: 'Another email send is already running.'};
  try {
    const hub = getHub_();
    const queue = hub.getSheetByName('Email Queue');
    const log = hub.getSheetByName('Outreach Log');
    const queueHeaders = requireHeaders_(queue, [
      'Queue ID', 'Entity Type', 'Entity ID', 'Recipient Name', 'Recipient Email', 'Template ID',
      'Subject', 'Body', 'Status', 'Approved?', 'Scheduled For', 'Provider Message ID', 'Sent At', 'Error',
      'Referral Owner', 'Sender Email', 'Delivery Provider'
    ]);
    const logHeaders = requireHeaders_(log, [
      'Outreach ID', 'Timestamp', 'Entity Type', 'Entity ID', 'Name', 'Channel', 'Recipient',
      'Campaign', 'Subject', 'Outcome', 'Sent By', 'Notes'
    ]);
    const values = queue.getDataRange().getValues();
    if (values.length < 2) return {ok: true, sent: 0, failed: 0, deferred: 0};

    const properties = PropertiesService.getScriptProperties();
    const runLimit = Math.max(1, Number(properties.getProperty('REMINDER_RUN_LIMIT')) || 75);
    const brevoDailyLimit = Math.max(1, Number(properties.getProperty('BREVO_DAILY_LIMIT')) || 300);
    let brevoSentToday = countBrevoSendsToday_(log, logHeaders);
    const gmailQuota = MailApp.getRemainingDailyQuota();
    let gmailSentThisRun = 0;
    let sent = 0;
    let failed = 0;
    let deferred = 0;

    for (let rowIndex = 1; rowIndex < values.length && sent < runLimit; rowIndex += 1) {
      const row = values[rowIndex];
      const approved = String(row[queueHeaders['Approved?']] || '').toLowerCase() === 'yes';
      const status = String(row[queueHeaders.Status] || '');
      const scheduled = row[queueHeaders['Scheduled For']];
      if (!approved || status !== 'Ready') continue;
      if (scheduled instanceof Date && scheduled.getTime() > Date.now()) continue;

      const sheetRow = rowIndex + 1;
      try {
        const recipient = normalizeEmail_(row[queueHeaders['Recipient Email']]);
        const subject = cleanText_(row[queueHeaders.Subject], 500);
        const body = cleanText_(row[queueHeaders.Body], 10000);
        const recipientName = cleanText_(row[queueHeaders['Recipient Name']], 160);
        const referralOwner = cleanText_(row[queueHeaders['Referral Owner']], 80);
        const configuredSender = getSenderForOwner_(referralOwner);
        const senderEmail = normalizeEmail_(row[queueHeaders['Sender Email']] || configuredSender.email);
        const provider = cleanText_(row[queueHeaders['Delivery Provider']] || configuredSender.provider, 40);
        if (!recipient || !subject || !body || !referralOwner || !senderEmail || !provider) {
          throw new Error('Recipient, subject, body, referral owner, sender email, and delivery provider are required.');
        }
        if (senderEmail !== configuredSender.email || provider.toLowerCase() !== configuredSender.provider.toLowerCase()) {
          throw new Error(`Queue sender routing does not match referral owner ${referralOwner}.`);
        }

        if (configuredSender.provider === 'Brevo' && brevoSentToday >= brevoDailyLimit) {
          queue.getRange(sheetRow, queueHeaders.Error + 1).setValue(`Deferred: Brevo daily limit of ${brevoDailyLimit} reached.`);
          deferred += 1;
          continue;
        }
        if (configuredSender.provider === 'Gmail' && gmailSentThisRun >= gmailQuota) {
          queue.getRange(sheetRow, queueHeaders.Error + 1).setValue('Deferred: Gmail daily recipient quota reached.');
          deferred += 1;
          continue;
        }

        const message = {
          recipient,
          recipientName,
          subject,
          body,
          senderEmail,
          senderName: configuredSender.name
        };
        const result = configuredSender.provider === 'Brevo'
          ? sendViaBrevo_(message)
          : sendViaGmail_(message);

        queue.getRange(sheetRow, queueHeaders.Status + 1).setValue('Sent');
        queue.getRange(sheetRow, queueHeaders['Provider Message ID'] + 1).setValue(result.messageId || '');
        queue.getRange(sheetRow, queueHeaders['Sent At'] + 1).setValue(new Date());
        queue.getRange(sheetRow, queueHeaders.Error + 1).clearContent();

        const logRow = new Array(log.getLastColumn()).fill('');
        logRow[logHeaders['Outreach ID']] = Utilities.getUuid();
        logRow[logHeaders.Timestamp] = new Date();
        logRow[logHeaders['Entity Type']] = row[queueHeaders['Entity Type']];
        logRow[logHeaders['Entity ID']] = row[queueHeaders['Entity ID']];
        logRow[logHeaders.Name] = recipientName;
        logRow[logHeaders.Channel] = 'Email';
        logRow[logHeaders.Recipient] = recipient;
        logRow[logHeaders.Campaign] = row[queueHeaders['Template ID']];
        logRow[logHeaders.Subject] = subject;
        logRow[logHeaders.Outcome] = 'Sent';
        logRow[logHeaders['Sent By']] = senderEmail;
        logRow[logHeaders.Notes] = `Provider: ${result.provider}; referral owner: ${referralOwner}; message ID: ${result.messageId || 'not returned'}.`;
        log.appendRow(logRow);
        if (result.provider === 'Brevo') brevoSentToday += 1;
        if (result.provider === 'Gmail') gmailSentThisRun += 1;
        sent += 1;
      } catch (error) {
        queue.getRange(sheetRow, queueHeaders.Status + 1).setValue('Failed');
        queue.getRange(sheetRow, queueHeaders.Error + 1).setValue(cleanText_(error && error.message ? error.message : error, 1000));
        failed += 1;
      }
    }
    return {
      ok: true,
      sent,
      failed,
      deferred,
      brevoSentToday,
      brevoDailyLimit,
      gmailQuotaAtStart: gmailQuota,
      gmailSentThisRun
    };
  } finally {
    lock.releaseLock();
  }
}

function syncTasksToCalendar() {
  const sheet = getHub_().getSheetByName('Tasks');
  const headers = requireHeaders_(sheet, [
    'Task', 'Status', 'Due Date', 'Due Time', 'Calendar Event ID', 'Workstream',
    'Related Entity Type', 'Related Entity ID', 'Notes'
  ]);
  const values = sheet.getDataRange().getValues();
  if (values.length < 2) return {ok: true, created: 0};
  const calendar = CalendarApp.getDefaultCalendar();
  let created = 0;

  for (let rowIndex = 1; rowIndex < values.length; rowIndex += 1) {
    const row = values[rowIndex];
    const title = cleanText_(row[headers.Task], 500);
    const status = String(row[headers.Status] || '');
    const dueDate = row[headers['Due Date']];
    const dueTime = row[headers['Due Time']];
    const existingEventId = row[headers['Calendar Event ID']];
    if (!title || !(dueDate instanceof Date) || existingEventId) continue;
    if (status === 'Done' || status === 'Cancelled') continue;

    const description = [
      `Workstream: ${cleanText_(row[headers.Workstream], 120)}`,
      `Related: ${cleanText_(row[headers['Related Entity Type']], 80)} ${cleanText_(row[headers['Related Entity ID']], 120)}`,
      cleanText_(row[headers.Notes], 1500)
    ].join('\n');
    let event;
    if (dueTime instanceof Date) {
      const start = new Date(dueDate);
      start.setHours(dueTime.getHours(), dueTime.getMinutes(), 0, 0);
      event = calendar.createEvent(title, start, new Date(start.getTime() + 60 * 60 * 1000), {description});
    } else {
      event = calendar.createAllDayEvent(title, dueDate, {description});
    }
    sheet.getRange(rowIndex + 1, headers['Calendar Event ID'] + 1).setValue(event.getId());
    created += 1;
  }
  return {ok: true, created};
}

function doPost(event) {
  const hub = getHub_();
  const sync = hub.getSheetByName('Fare Sync');
  const headers = requireHeaders_(sync, [
    'Sync ID', 'Run At', 'Direction', 'Entity Type', 'Entity ID', 'Fare Record ID',
    'Operation', 'Status', 'HTTP Status', 'Changed Fields', 'Source Updated At',
    'Google Updated At', 'Error', 'Retry?', 'Notes'
  ]);
  const expected = PropertiesService.getScriptProperties().getProperty('FARE_WEBHOOK_SECRET');
  const supplied = event && event.parameter ? String(event.parameter.token || '') : '';
  if (!expected || supplied !== expected) {
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'unauthorized'}))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const row = new Array(sync.getLastColumn()).fill('');
  row[headers['Sync ID']] = Utilities.getUuid();
  row[headers['Run At']] = new Date();
  row[headers.Direction] = 'Fare -> Google';
  try {
    const payload = JSON.parse(event.postData && event.postData.contents ? event.postData.contents : '{}');
    const driverId = cleanText_(payload.fare_driver_id || payload.id, 160);
    row[headers['Entity Type']] = cleanText_(payload.entity_type || 'Driver', 80);
    row[headers['Entity ID']] = driverId;
    row[headers['Fare Record ID']] = driverId;
    row[headers.Operation] = 'Webhook received';
    row[headers.Status] = 'Received';
    row[headers['HTTP Status']] = 200;
    row[headers['Changed Fields']] = cleanText_(JSON.stringify(payload), 45000);
    row[headers['Source Updated At']] = cleanText_(payload.updated_at, 100);
    row[headers['Google Updated At']] = new Date();
    row[headers['Retry?']] = 'No';
    row[headers.Notes] = 'Logging-only mode. Awaiting Fare-approved field mapping.';
    sync.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ok: true}))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    row[headers['Entity Type']] = 'Unknown';
    row[headers.Operation] = 'Webhook error';
    row[headers.Status] = 'Failed';
    row[headers['HTTP Status']] = 400;
    row[headers['Google Updated At']] = new Date();
    row[headers.Error] = cleanText_(error && error.message ? error.message : error, 1000);
    row[headers['Retry?']] = 'Yes';
    sync.appendRow(row);
    return ContentService.createTextOutput(JSON.stringify({ok: false, error: 'invalid payload'}))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
