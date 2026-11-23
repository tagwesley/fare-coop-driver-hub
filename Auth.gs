function requestLoginCode(email) {
  const normalizedEmail = normalizeEmail_(email);
  if (!/^\S+@\S+\.\S+$/.test(normalizedEmail) || normalizedEmail.length > 254) {
    return {ok: true, message: 'If that email is on the Fare Co-op driver roster, a sign-in code is on the way.'};
  }

  const cache = CacheService.getScriptCache();
  const emailKey = sha256Hex_(normalizedEmail);
  const throttleKey = `otp-throttle:${emailKey}`;
  if (cache.get(throttleKey)) {
    return {ok: true, message: 'If that email is on the Fare Co-op driver roster, a sign-in code is on the way.'};
  }
  cache.put(throttleKey, '1', 60);

  const driver = findDriverByEmail_(normalizedEmail);
  if (!driver) {
    return {ok: true, message: 'If that email is on the Fare Co-op driver roster, a sign-in code is on the way.'};
  }

  if (MailApp.getRemainingDailyQuota() < 1) {
    return {ok: false, message: 'Sign-in email is temporarily unavailable. Please contact Fare Co-op support.'};
  }

  const code = createSixDigitCode_();
  const salt = Utilities.getUuid();
  const record = {
    email: normalizedEmail,
    driverId: driver.id,
    codeHash: sha256Hex_(`${salt}:${code}`),
    salt,
    attempts: 0,
    createdAt: Date.now()
  };
  cache.put(`otp:${emailKey}`, JSON.stringify(record), PORTAL_OTP_SECONDS);

  const firstName = cleanText_(driver.name, 120).split(/\s+/)[0] || 'driver';
  const subject = 'Your Fare Co-op Driver Hub sign-in code';
  const plainBody = [
    `Hi ${firstName},`,
    '',
    `Your Fare Co-op Driver Hub code is: ${code}`,
    '',
    'It expires in 10 minutes. If you did not request this code, you can ignore this email.',
    '',
    'Fare Co-op Driver Operations'
  ].join('\n');

  MailApp.sendEmail({
    to: normalizedEmail,
    subject,
    body: plainBody,
    htmlBody: `<div style="font-family:Arial,sans-serif;color:#17243a;line-height:1.5"><p>Hi ${htmlEscapeServer_(firstName)},</p><p>Your Fare Co-op Driver Hub code is:</p><p style="font-size:30px;font-weight:800;letter-spacing:8px;margin:20px 0">${code}</p><p>It expires in 10 minutes. If you did not request this code, you can ignore this email.</p><p>Fare Co-op Driver Operations</p></div>`,
    name: 'Fare Co-op Driver Hub'
  });

  return {ok: true, message: 'If that email is on the Fare Co-op driver roster, a sign-in code is on the way.'};
}

function verifyLoginCode(email, code) {
  const normalizedEmail = normalizeEmail_(email);
  const normalizedCode = String(code || '').replace(/\D/g, '').slice(0, 6);
  if (!/^\S+@\S+\.\S+$/.test(normalizedEmail) || !/^\d{6}$/.test(normalizedCode)) {
    return {ok: false, message: 'Enter the six-digit code from your email.'};
  }

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const cache = CacheService.getScriptCache();
    const emailKey = sha256Hex_(normalizedEmail);
    const otpKey = `otp:${emailKey}`;
    const rawRecord = cache.get(otpKey);
    if (!rawRecord) return {ok: false, message: 'That code expired. Request a new one.'};

    const record = JSON.parse(rawRecord);
    record.attempts = Number(record.attempts || 0) + 1;
    if (record.attempts > 5) {
      cache.remove(otpKey);
      return {ok: false, message: 'Too many attempts. Request a new code.'};
    }

    const expectedHash = sha256Hex_(`${record.salt}:${normalizedCode}`);
    if (expectedHash !== record.codeHash || normalizedEmail !== record.email) {
      cache.put(otpKey, JSON.stringify(record), Math.max(60, PORTAL_OTP_SECONDS - Math.floor((Date.now() - record.createdAt) / 1000)));
      return {ok: false, message: 'That code is not correct.'};
    }

    const driver = findDriverByEmail_(normalizedEmail);
    if (!driver || normalizeId_(driver.id) !== normalizeId_(record.driverId)) {
      cache.remove(otpKey);
      return {ok: false, message: 'Your driver record is not available. Contact Fare Co-op support.'};
    }

    cache.remove(otpKey);
    const token = `${Utilities.getUuid()}${Utilities.getUuid()}`.replace(/-/g, '');
    const session = {
      driverId: normalizeId_(driver.id),
      email: normalizedEmail,
      createdAt: nowIso_()
    };
    cache.put(`session:${sha256Hex_(token)}`, JSON.stringify(session), PORTAL_SESSION_SECONDS);
    return {ok: true, token, expiresSeconds: PORTAL_SESSION_SECONDS};
  } finally {
    lock.releaseLock();
  }
}

function logoutDriver(token) {
  const cleanToken = String(token || '').trim();
  if (cleanToken) CacheService.getScriptCache().remove(`session:${sha256Hex_(cleanToken)}`);
  return {ok: true};
}

function getSession_(token) {
  const cleanToken = String(token || '').trim();
  if (!/^[a-f0-9]{64}$/i.test(cleanToken)) throw new Error('AUTH_REQUIRED');
  const cache = CacheService.getScriptCache();
  const key = `session:${sha256Hex_(cleanToken)}`;
  const rawSession = cache.get(key);
  if (!rawSession) throw new Error('AUTH_REQUIRED');
  const session = JSON.parse(rawSession);
  const driver = findDriverByEmail_(session.email);
  if (!driver || normalizeId_(driver.id) !== normalizeId_(session.driverId)) {
    cache.remove(key);
    throw new Error('AUTH_REQUIRED');
  }
  cache.put(key, rawSession, PORTAL_SESSION_SECONDS);
  session.driver = driver;
  return session;
}

function findDriverByEmail_(email) {
  const sheet = getHub_().getSheetByName('Driver Index');
  const headers = requireHeaders_(sheet, ['Driver ID', 'Driver Name', 'Email']);
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return null;
  const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getDisplayValues();
  const target = normalizeEmail_(email);
  for (let index = 0; index < values.length; index += 1) {
    const row = values[index];
    if (normalizeEmail_(row[headers.Email]) === target) {
      return {
        rowNumber: index + 2,
        id: normalizeId_(row[headers['Driver ID']]),
        name: cleanText_(row[headers['Driver Name']], 160),
        email: target,
        values: row,
        headers
      };
    }
  }
  return null;
}

function createSixDigitCode_() {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, Utilities.getUuid(), Utilities.Charset.UTF_8);
  let value = 0;
  for (let index = 0; index < 6; index += 1) value = (value * 256 + ((bytes[index] + 256) % 256)) % 1000000;
  return String(value).padStart(6, '0');
}

function htmlEscapeServer_(value) {
  return String(value || '')
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
