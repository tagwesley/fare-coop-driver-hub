function getPortalData(token) {
  const session = getSession_(token);
  const driver = buildDriverProfile_(session.driver);
  const action = getDriverAction_(driver.id);
  return {
    ok: true,
    generatedAt: formatDateTime_(new Date()),
    driver,
    availability: action.availability,
    operationalStatus: action.operationalStatus,
    dispatches: getDriverDispatches_(driver.id),
    events: getDriverEvents_(driver.city),
    links: {
      fareDriverPortal: 'https://www.fare.coop/drivers/',
      membership: 'https://www.fare.coop/membership/',
      tutorials: 'https://www.fare.coop/tutorials/',
      support: 'https://www.fare.coop/support-tickets/'
    }
  };
}

function setAvailability(token, availability) {
  const session = getSession_(token);
  const value = cleanText_(availability, 20);
  if (!['Yes', 'No', 'Unknown'].includes(value)) throw new Error('Invalid availability value.');

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getHub_().getSheetByName('Driver Actions');
    const headers = requireHeaders_(sheet, [
      'Action ID', 'Driver ID', 'Owner', 'Operational Status', 'Available for Dispatch?',
      'Market', 'Priority', 'Updated At', 'Updated By'
    ]);
    const driverId = normalizeId_(session.driverId);
    const lastRow = sheet.getLastRow();
    let rowNumber = 0;
    if (lastRow >= 2) {
      const ids = sheet.getRange(2, headers['Driver ID'] + 1, lastRow - 1, 1).getDisplayValues();
      const matchIndex = ids.findIndex(row => normalizeId_(row[0]) === driverId);
      if (matchIndex >= 0) rowNumber = matchIndex + 2;
    }

    const status = value === 'Yes' ? 'Open' : value === 'No' ? 'Paused' : 'Waiting';
    if (!rowNumber) {
      const row = new Array(sheet.getLastColumn()).fill('');
      row[headers['Action ID']] = Utilities.getUuid();
      row[headers['Driver ID']] = driverId;
      row[headers.Owner] = session.email;
      row[headers['Operational Status']] = status;
      row[headers['Available for Dispatch?']] = value;
      row[headers.Market] = cleanText_(session.driver.values[session.driver.headers.City], 160);
      row[headers.Priority] = 'Normal';
      row[headers['Updated At']] = new Date();
      row[headers['Updated By']] = `Driver Portal: ${session.email}`;
      sheet.appendRow(row);
    } else {
      sheet.getRange(rowNumber, headers['Operational Status'] + 1).setValue(status);
      sheet.getRange(rowNumber, headers['Available for Dispatch?'] + 1).setValue(value);
      sheet.getRange(rowNumber, headers['Updated At'] + 1).setValue(new Date());
      sheet.getRange(rowNumber, headers['Updated By'] + 1).setValue(`Driver Portal: ${session.email}`);
    }
    SpreadsheetApp.flush();
    return {ok: true, availability: value, operationalStatus: status, message: 'Availability updated.'};
  } finally {
    lock.releaseLock();
  }
}

function updateDispatchStatus(token, dispatchId, requestedStatus) {
  const session = getSession_(token);
  const id = normalizeId_(dispatchId);
  const status = cleanText_(requestedStatus, 30);
  const allowed = ['Confirmed', 'En Route', 'Completed', 'Cancelled'];
  if (!id || !allowed.includes(status)) throw new Error('Invalid dispatch update.');

  const transitions = {
    Unassigned: [],
    Assigned: ['Confirmed', 'Cancelled'],
    Confirmed: ['En Route', 'Cancelled'],
    'En Route': ['Completed', 'Cancelled'],
    Completed: [],
    Cancelled: []
  };

  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getHub_().getSheetByName('Dispatch');
    const headers = requireHeaders_(sheet, [
      'Dispatch ID', 'Driver ID', 'Assignment Status', 'Confirmed At', 'Completed At', 'Notes'
    ]);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) throw new Error('Dispatch not found.');
    const values = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const matchIndex = values.findIndex(row => normalizeId_(row[headers['Dispatch ID']]) === id);
    if (matchIndex < 0) throw new Error('Dispatch not found.');
    const row = values[matchIndex];
    if (normalizeId_(row[headers['Driver ID']]) !== normalizeId_(session.driverId)) {
      throw new Error('You cannot update this dispatch.');
    }

    const currentStatus = cleanText_(row[headers['Assignment Status']] || 'Unassigned', 30);
    if (!(transitions[currentStatus] || []).includes(status)) {
      throw new Error(`This dispatch cannot move from ${currentStatus} to ${status}. Refresh the app and try again.`);
    }

    const rowNumber = matchIndex + 2;
    const now = new Date();
    sheet.getRange(rowNumber, headers['Assignment Status'] + 1).setValue(status);
    if (status === 'Confirmed' && !row[headers['Confirmed At']]) {
      sheet.getRange(rowNumber, headers['Confirmed At'] + 1).setValue(now);
    }
    if (status === 'Completed' && !row[headers['Completed At']]) {
      sheet.getRange(rowNumber, headers['Completed At'] + 1).setValue(now);
    }
    const existingNotes = cleanText_(row[headers.Notes], 2500);
    const portalNote = `[${formatDateTime_(now)}] Driver portal update: ${status}.`;
    sheet.getRange(rowNumber, headers.Notes + 1).setValue(existingNotes ? `${existingNotes}\n${portalNote}` : portalNote);
    SpreadsheetApp.flush();
    return {ok: true, status, message: 'Dispatch status updated.'};
  } finally {
    lock.releaseLock();
  }
}

function submitHelpRequest(token, category, details) {
  const session = getSession_(token);
  const cleanCategory = cleanText_(category, 40);
  const cleanDetails = cleanText_(details, 1000);
  const allowedCategories = ['Onboarding', 'Documents', 'Vehicle', 'Payments', 'Dispatch', 'App issue', 'Other'];
  if (!allowedCategories.includes(cleanCategory)) throw new Error('Choose a support category.');
  if (cleanDetails.length < 5) throw new Error('Add a few details so the team can help.');

  const sheet = getHub_().getSheetByName('Tasks');
  const headers = requireHeaders_(sheet, [
    'Task ID', 'Created', 'Owner', 'Workstream', 'Related Entity Type', 'Related Entity ID',
    'Task', 'Priority', 'Status', 'Notes'
  ]);
  const row = new Array(sheet.getLastColumn()).fill('');
  row[headers['Task ID']] = Utilities.getUuid();
  row[headers.Created] = new Date();
  row[headers.Owner] = PORTAL_SUPPORT_OWNER;
  row[headers.Workstream] = 'Driver Support';
  row[headers['Related Entity Type']] = 'Driver';
  row[headers['Related Entity ID']] = session.driverId;
  row[headers.Task] = `${cleanCategory}: ${session.driver.name}`;
  row[headers.Priority] = cleanCategory === 'Dispatch' ? 'High' : 'Normal';
  row[headers.Status] = 'Not Started';
  row[headers.Notes] = `${cleanDetails}\nSubmitted by ${session.driver.name} (${session.driverId}) through the Driver Hub.`;
  sheet.appendRow(row);
  return {ok: true, message: 'Your request was sent to Fare Co-op operations.'};
}

function expressEventInterest(token, eventId) {
  const session = getSession_(token);
  const id = normalizeId_(eventId);
  if (!id) throw new Error('Event not found.');

  const hub = getHub_();
  const eventsSheet = hub.getSheetByName('Events');
  const eventHeaders = requireHeaders_(eventsSheet, ['Event ID', 'Event Name', 'Start', 'Venue', 'City / Market']);
  const eventValues = eventsSheet.getDataRange().getValues();
  const eventIndex = eventValues.slice(1).findIndex(row => normalizeId_(row[eventHeaders['Event ID']]) === id);
  if (eventIndex < 0) throw new Error('Event not found.');
  const event = eventValues[eventIndex + 1];

  const tasksSheet = hub.getSheetByName('Tasks');
  const taskHeaders = requireHeaders_(tasksSheet, [
    'Task ID', 'Created', 'Owner', 'Workstream', 'Related Entity Type', 'Related Entity ID',
    'Task', 'Priority', 'Status', 'Due Date', 'Notes'
  ]);
  const taskValues = tasksSheet.getDataRange().getDisplayValues();
  const duplicate = taskValues.slice(1).some(row =>
    row[taskHeaders.Workstream] === 'Event Staffing' &&
    normalizeId_(row[taskHeaders['Related Entity ID']]) === id &&
    String(row[taskHeaders.Notes] || '').includes(`Driver ID: ${session.driverId}`)
  );
  if (duplicate) return {ok: true, message: 'The operations team already has your interest for this event.'};

  const row = new Array(tasksSheet.getLastColumn()).fill('');
  row[taskHeaders['Task ID']] = Utilities.getUuid();
  row[taskHeaders.Created] = new Date();
  row[taskHeaders.Owner] = PORTAL_SUPPORT_OWNER;
  row[taskHeaders.Workstream] = 'Event Staffing';
  row[taskHeaders['Related Entity Type']] = 'Event';
  row[taskHeaders['Related Entity ID']] = id;
  row[taskHeaders.Task] = `Staffing interest: ${session.driver.name} — ${cleanText_(event[eventHeaders['Event Name']], 160)}`;
  row[taskHeaders.Priority] = 'Normal';
  row[taskHeaders.Status] = 'Not Started';
  if (event[eventHeaders.Start] instanceof Date) row[taskHeaders['Due Date']] = event[eventHeaders.Start];
  row[taskHeaders.Notes] = [
    `Driver ID: ${session.driverId}`,
    `Driver: ${session.driver.name}`,
    `Email: ${session.email}`,
    `Venue: ${cleanText_(event[eventHeaders.Venue], 200)}`,
    `Market: ${cleanText_(event[eventHeaders['City / Market']], 160)}`,
    'Submitted through the Driver Hub.'
  ].join('\n');
  tasksSheet.appendRow(row);
  return {ok: true, message: 'Your interest was sent to the event staffing team.'};
}

function buildDriverProfile_(driverRecord) {
  const row = driverRecord.values;
  const h = driverRecord.headers;
  return {
    id: driverRecord.id,
    name: driverRecord.name,
    email: driverRecord.email,
    phone: cleanText_(row[h.Phone], 80),
    city: cleanText_(row[h.City], 160),
    source: cleanText_(row[h.Source], 80),
    referralOwner: cleanText_(row[h['Referral Owner']], 80),
    stage: cleanText_(row[h.Stage], 120),
    stageProgress: cleanText_(row[h['Stage Progress']], 30),
    checklistProgress: cleanText_(row[h['Checklist Progress']], 30),
    nextRequiredStep: cleanText_(row[h['Next Required Step']], 300) || 'Check your Fare Co-op driver app for the next step.',
    background: cleanText_(row[h['Background / Agreement / Quiz']], 40),
    documents: cleanText_(row[h['Documents Uploaded']], 40),
    vehicle: cleanText_(row[h['Vehicle Added']], 40),
    stripe: cleanText_(row[h['Stripe Completed']], 40),
    membershipType: cleanText_(row[h['Membership Type']], 100),
    updatedAt: cleanText_(row[h['Updated At']], 100)
  };
}

function getDriverAction_(driverId) {
  const sheet = getHub_().getSheetByName('Driver Actions');
  const headers = requireHeaders_(sheet, ['Driver ID', 'Operational Status', 'Available for Dispatch?']);
  const values = sheet.getDataRange().getDisplayValues();
  const match = values.slice(1).find(row => normalizeId_(row[headers['Driver ID']]) === normalizeId_(driverId));
  return {
    availability: match ? cleanText_(match[headers['Available for Dispatch?']], 20) || 'Unknown' : 'Unknown',
    operationalStatus: match ? cleanText_(match[headers['Operational Status']], 30) || 'Waiting' : 'Waiting'
  };
}

function getDriverDispatches_(driverId) {
  const sheet = getHub_().getSheetByName('Dispatch');
  const headers = requireHeaders_(sheet, [
    'Dispatch ID', 'Ride / Event ID', 'Rider / Client', 'Pickup', 'Dropoff', 'Pickup Time',
    'Driver ID', 'Assignment Status', 'Confirmed At', 'Completed At', 'Fare Ride ID', 'Notes'
  ]);
  const values = sheet.getDataRange().getValues();
  return values.slice(1)
    .filter(row => normalizeId_(row[headers['Driver ID']]) === normalizeId_(driverId))
    .map(row => ({
      id: normalizeId_(row[headers['Dispatch ID']]),
      rideEventId: cleanText_(row[headers['Ride / Event ID']], 100),
      riderClient: cleanText_(row[headers['Rider / Client']], 160),
      pickup: cleanText_(row[headers.Pickup], 300),
      dropoff: cleanText_(row[headers.Dropoff], 300),
      pickupTime: formatDateTime_(row[headers['Pickup Time']]),
      pickupTimestamp: row[headers['Pickup Time']] instanceof Date ? row[headers['Pickup Time']].getTime() : 0,
      status: cleanText_(row[headers['Assignment Status']], 30) || 'Assigned',
      confirmedAt: formatDateTime_(row[headers['Confirmed At']]),
      completedAt: formatDateTime_(row[headers['Completed At']]),
      fareRideId: cleanText_(row[headers['Fare Ride ID']], 100),
      notes: cleanText_(row[headers.Notes], 800)
    }))
    .sort((a, b) => (b.pickupTimestamp || 0) - (a.pickupTimestamp || 0))
    .slice(0, 30);
}

function getDriverEvents_(driverCity) {
  const sheet = getHub_().getSheetByName('Events');
  const headers = requireHeaders_(sheet, [
    'Event ID', 'Event Name', 'Venue', 'Address', 'Start', 'End', 'City / Market',
    'Staffing Needed', 'Drivers Needed', 'Event Status', 'Ticketing URL', 'Notes'
  ]);
  const values = sheet.getDataRange().getValues();
  const city = String(driverCity || '').toLowerCase();
  const yesterday = new Date();
  yesterday.setHours(0, 0, 0, 0);
  yesterday.setDate(yesterday.getDate() - 1);

  return values.slice(1)
    .filter(row => normalizeId_(row[headers['Event ID']]))
    .filter(row => {
      const status = String(row[headers['Event Status']] || '').toLowerCase();
      if (['completed', 'cancelled'].includes(status)) return false;
      const start = row[headers.Start];
      if (start instanceof Date && start.getTime() < yesterday.getTime()) return false;
      const market = String(row[headers['City / Market']] || '').toLowerCase().trim();
      return !market || !city || city.includes(market) || market.includes(city.split(',')[0].trim());
    })
    .map(row => ({
      id: normalizeId_(row[headers['Event ID']]),
      name: cleanText_(row[headers['Event Name']], 180),
      venue: cleanText_(row[headers.Venue], 180),
      address: cleanText_(row[headers.Address], 300),
      start: formatDateTime_(row[headers.Start]),
      startTimestamp: row[headers.Start] instanceof Date ? row[headers.Start].getTime() : 0,
      end: formatDateTime_(row[headers.End]),
      market: cleanText_(row[headers['City / Market']], 160),
      staffingNeeded: cleanText_(row[headers['Staffing Needed']], 40),
      driversNeeded: cleanText_(row[headers['Drivers Needed']], 40),
      status: cleanText_(row[headers['Event Status']], 40),
      ticketingUrl: /^https:\/\//i.test(String(row[headers['Ticketing URL']] || '')) ? String(row[headers['Ticketing URL']]) : '',
      notes: cleanText_(row[headers.Notes], 500)
    }))
    .sort((a, b) => (a.startTimestamp || Number.MAX_SAFE_INTEGER) - (b.startTimestamp || Number.MAX_SAFE_INTEGER))
    .slice(0, 20);
}
