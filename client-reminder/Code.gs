/**
 * Client Follow-Up Reminder System v2.0
 * 
 * Relationship intelligence and reminder layer for the Gmail Automation Suite.
 * Reads from email-to-spreadsheet Logger, builds thread state and client pulse,
 * and generates actionable reminders.
 * 
 * Setup: Click Extensions → Apps Script, paste this code, save, then run setupTemplate().
 */

// Configuration Constants
const SHEET_NAMES = {
  SETTINGS: 'Settings',
  CLIENTS: 'Clients',
  INTERACTIONS: 'Interactions',
  THREAD_STATE: 'ThreadState',
  CLIENT_PULSE: 'ClientPulse',
  REMINDERS: 'Reminders',
  UNMATCHED_CONTACTS: 'UnmatchedContacts',
  SYNC_LOG: 'SyncLog'
};

// Default thresholds
const DEFAULT_THRESHOLDS = {
  typicalSilenceDays: 14,      // Expected days between contacts
  silenceWarningDays: 21,      // Days before "Warning" status
  silenceCriticalDays: 30,     // Days before "Critical" status
  replyTimeWindow: 7,          // Days to consider for avg reply time
  reminderCooldown: 3          // Days before repeating same reminder
};

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('📋 Client Reminder')
    .addItem('⚙️ Setup Template', 'setupTemplate')
    .addSeparator()
    .addItem('🔄 Full Sync', 'runFullSync')
    .addItem('📊 Rebuild Pulse', 'rebuildClientPulse')
    .addItem('🔔 Generate Reminders', 'generateReminders')
    .addSeparator()
    .addItem('📋 Check Status', 'showStatus')
    .addItem('📊 Show Unmatched', 'showUnmatchedCount')
    .addToUi();
}

/**
 * One-click setup: Creates all sheets with headers, formatting, and sample data
 */
function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  
  createSettingsSheet(ss);
  createClientsSheet(ss);
  createInteractionsSheet(ss);
  createThreadStateSheet(ss);
  createClientPulseSheet(ss);
  createRemindersSheet(ss);
  createUnmatchedContactsSheet(ss);
  createSyncLogSheet(ss);
  
  ui.alert('✅ Setup Complete', 
    'Template created!\n\n' +
    'Next steps:\n' +
    '1. Add source spreadsheet ID in Settings\n' +
    '2. Add your clients in the Clients tab\n' +
    '3. Run "🔄 Full Sync" to import interactions\n' +
    '4. Set up daily trigger for automatic reminders', 
    ui.ButtonSet.OK);
}

// ============================================================================
// SHEET CREATION FUNCTIONS
// ============================================================================

function createSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  
  const settings = [
    ['Setting', 'Value', 'Description'],
    ['Source Spreadsheet ID', '', 'ID of the email-to-spreadsheet Logger'],
    ['Source Log Sheet Name', 'Log', 'Name of the log sheet in source'],
    ['Auto Sync Enabled', 'Yes', 'Enable automatic sync'],
    ['Sync Interval (hours)', '6', 'Hours between automatic syncs'],
    ['Typical Silence Days', '14', 'Expected days between client contacts'],
    ['Silence Warning Days', '21', 'Days before warning status'],
    ['Silence Critical Days', '30', 'Days before critical status'],
    ['Reply Time Window (days)', '7', 'Days to consider for avg reply time'],
    ['Reminder Cooldown (days)', '3', 'Days before repeating same reminder'],
    ['Internal Emails', '', 'Your email addresses (comma-separated)'],
    ['Last Sync', 'Never', 'Timestamp of last successful sync'],
    ['Last Sync Count', '0', 'Number of interactions synced last run']
  ];
  
  sheet.getRange(1, 1, settings.length, 3).setValues(settings);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.getRange('A:A').setFontWeight('bold');
  sheet.setColumnWidth(1, 25);
  sheet.setColumnWidth(2, 40);
  sheet.setColumnWidth(3, 50);
}

function createClientsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.CLIENTS);
  
  const headers = [
    'Client ID',           // A - Unique ID (e.g., C001)
    'Client Name',         // B - Display name
    'Primary Email',       // C - Main contact email
    'Secondary Emails',    // D - Additional emails (comma-separated)
    'Company',             // E - Company name
    'Status',              // F - Active / Inactive / Prospect
    'Priority',            // G - High / Medium / Low
    'First Contact',       // H - Date of first interaction
    'Last Contact',        // I - Date of most recent interaction
    'Total Interactions',  // J - Count of all interactions
    'Notes',               // K - Free-form notes
    'Created At',          // L - When this client was added
    'Updated At'           // M - When this client was last modified
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
  
  // Sample data
  const sampleData = [
    ['C001', 'Acme Corp', 'contact@acme.com', 'support@acme.com', 'Acme Corporation', 'Active', 'High', '', '', '', '', '', ''],
    ['C002', 'Globex Inc', 'info@globex.com', '', 'Globex Industries', 'Active', 'Medium', '', '', '', '', '', '']
  ];
  sheet.getRange(2, 1, sampleData.length, sampleData[0].length).setValues(sampleData);
  
  // Conditional formatting for Status
  const activeRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Active')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('F2:F')])
    .build();
  const inactiveRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Inactive')
    .setBackground('#fce5cd')
    .setRanges([sheet.getRange('F2:F')])
    .build();
  sheet.setConditionalFormatRules([activeRule, inactiveRule]);
}

function createInteractionsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.INTERACTIONS);
  
  const headers = [
    'Interaction ID',      // A - Unique key: ClientID__MessageID
    'Client ID',           // B - Reference to Clients
    'Message ID',          // C - Gmail message ID
    'Thread ID',           // D - Gmail thread ID
    'Date',                // E - When the email was sent
    'Direction',           // F - Inbound / Outbound
    'Primary Contact',     // G - Email address
    'Subject',             // H - Email subject
    'Snippet',             // I - Body preview
    'Gmail Link',          // J - Direct link
    'Synced At'            // K - When we imported this
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
}

function createThreadStateSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.THREAD_STATE);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.THREAD_STATE);
  
  const headers = [
    'Thread ID',           // A - Gmail thread ID
    'Client ID',           // B - Matched client
    'Message Count',       // C - Total messages in thread
    'First Message Date',  // D - Earliest message
    'Last Message Date',   // E - Most recent message
    'Last Sender',         // F - Who sent last message
    'Waiting On',          // G - 'Us' or 'Them'
    'Avg Reply Time (hrs)',// H - Average time between replies
    'Status',              // I - Active / Stale / Waiting
    'Risk Score',          // J - 0-100 risk score
    'Subject',             // K - First message subject
    'Gmail Link',          // L - Direct link
    'Updated At'           // M - When this was computed
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
}

function createClientPulseSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CLIENT_PULSE);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.CLIENT_PULSE);
  
  const headers = [
    'Client ID',              // A - Reference to Clients
    'Client Name',            // B - Display name
    'Last Interaction',       // C - Most recent interaction date
    'Last Inbound',           // D - Most recent inbound email
    'Last Outbound',          // E - Most recent outbound email
    'Typical Silence Days',   // F - Expected days between contacts
    'Current Silence Days',   // G - Days since last contact
    'Silence Drift Score',    // H - (Current - Typical) / Typical
    'Adjusted Drift Score',   // I - Drift adjusted for client priority
    'Relationship Temperature',// J - Hot / Warm / Cool / Cold
    'Relationship Status',    // K - Healthy / Warning / Critical / Overdue
    'Waiting On',             // L - 'Us' or 'Them'
    'Active Threads',         // M - Number of active threads
    'Total Interactions',     // N - All-time interaction count
    'Reminder Score',         // O - 0-100 priority for reminder
    'Recommended Action',     // P - Suggested follow-up action
    'Updated At'              // Q - When this was computed
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
}

function createRemindersSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.REMINDERS);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.REMINDERS);
  
  const headers = [
    'Reminder ID',         // A - Unique ID
    'Client ID',           // B - Reference to Clients
    'Client Name',         // C - Display name
    'Type',                // D - FollowUp / Stale / Waiting / Critical
    'Priority',            // E - High / Medium / Low
    'Message',             // F - Human-readable reminder
    'Days Since Contact',  // G - How long since last contact
    'Thread ID',           // H - Related thread (if any)
    'Gmail Link',          // I - Direct link to thread
    'Status',              // J - Pending / Completed / Dismissed
    'Created At',          // K - When reminder was generated
    'Completed At'         // L - When marked complete
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
  
  // Conditional formatting for Status
  const pendingRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Pending')
    .setBackground('#fff2cc')
    .setRanges([sheet.getRange('J2:J')])
    .build();
  const completedRule = SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo('Completed')
    .setBackground('#d9ead3')
    .setRanges([sheet.getRange('J2:J')])
    .build();
  sheet.setConditionalFormatRules([pendingRule, completedRule]);
}

function createUnmatchedContactsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.UNMATCHED_CONTACTS);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.UNMATCHED_CONTACTS);
  
  const headers = [
    'Email',               // A - Unmatched email address
    'Domain',              // B - Email domain
    'First Seen',          // C - When first encountered
    'Last Seen',           // D - Most recent encounter
    'Interaction Count',   // E - How many times seen
    'Sample Subjects',     // F - Subject lines (pipe-separated)
    'Suggested Client',    // G - Auto-suggested client name
    'Status',              // H - New / Ignored / Added
    'Created At'           // I - When added to this list
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
}

function createSyncLogSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SYNC_LOG);
  if (sheet) return;
  
  sheet = ss.insertSheet(SHEET_NAMES.SYNC_LOG);
  
  const headers = [
    'Sync ID',             // A - Unique ID
    'Started At',          // B - When sync started
    'Completed At',        // C - When sync finished
    'Source Rows',         // D - Total rows in source
    'New Interactions',    // E - Newly synced interactions
    'Updated Threads',     // F - Threads updated
    'Updated Clients',     // G - Client pulses updated
    'New Reminders',       // H - Reminders generated
    'Status',              // I - Success / Partial / Failed
    'Error Message'        // J - If failed
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#4285f4').setFontColor('white');
  sheet.setFrozenRows(1);
}

// ============================================================================
// SETTINGS HELPERS
// ============================================================================

function getSettings(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return DEFAULT_THRESHOLDS;
  
  const data = sheet.getRange('A2:B' + sheet.getLastRow()).getValues();
  const settings = {};
  
  data.forEach(row => {
    if (row[0] && row[1]) {
      settings[row[0]] = row[1];
    }
  });
  
  return {
    sourceSpreadsheetId: settings['Source Spreadsheet ID'] || '',
    sourceLogSheetName: settings['Source Log Sheet Name'] || 'Log',
    autoSyncEnabled: settings['Auto Sync Enabled'] === 'Yes',
    syncIntervalHours: parseInt(settings['Sync Interval (hours)']) || 6,
    typicalSilenceDays: parseInt(settings['Typical Silence Days']) || DEFAULT_THRESHOLDS.typicalSilenceDays,
    silenceWarningDays: parseInt(settings['Silence Warning Days']) || DEFAULT_THRESHOLDS.silenceWarningDays,
    silenceCriticalDays: parseInt(settings['Silence Critical Days']) || DEFAULT_THRESHOLDS.silenceCriticalDays,
    replyTimeWindow: parseInt(settings['Reply Time Window (days)']) || DEFAULT_THRESHOLDS.replyTimeWindow,
    reminderCooldown: parseInt(settings['Reminder Cooldown (days)']) || DEFAULT_THRESHOLDS.reminderCooldown,
    internalEmails: (settings['Internal Emails'] || '').toString().split(',').map(e => e.trim().toLowerCase()).filter(e => e)
  };
}

function updateSetting(ss, key, value) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) return;
  
  const data = sheet.getRange('A2:A' + sheet.getLastRow()).getValues();
  for (let i = 0; i < data.length; i++) {
    if (data[i][0] === key) {
      sheet.getRange(i + 2, 2).setValue(value);
      return;
    }
  }
}

// ============================================================================
// CLIENT MATCHING
// ============================================================================

/**
 * Loads all clients into a lookup map
 * @returns {Object} - Map of email -> client object
 */
function loadClientLookup(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  if (!sheet || sheet.getLastRow() < 2) return {};
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 13).getValues();
  const lookup = {};
  
  data.forEach(row => {
    const client = {
      clientId: row[0],
      clientName: row[1],
      primaryEmail: (row[2] || '').toLowerCase(),
      secondaryEmails: (row[3] || '').toString().split(',').map(e => e.trim().toLowerCase()).filter(e => e),
      company: row[4],
      status: row[5],
      priority: row[6]
    };
    
    // Index by primary email
    if (client.primaryEmail) {
      lookup[client.primaryEmail] = client;
    }
    
    // Index by secondary emails
    client.secondaryEmails.forEach(email => {
      lookup[email] = client;
    });
  });
  
  return lookup;
}

/**
 * Matches an email address to a client
 * @param {string} email - Email address to match
 * @param {Object} lookup - Client lookup map
 * @returns {Object|null} - Matched client or null
 */
function matchClient(email, lookup) {
  if (!email) return null;
  return lookup[email.toLowerCase()] || null;
}

// ============================================================================
// PHASE 1: SOURCE SYNC
// ============================================================================

/**
 * Main sync function - imports data from source logger spreadsheet
 */
function runFullSync() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();
  const settings = getSettings(ss);
  
  if (!settings.sourceSpreadsheetId) {
    ui.alert('Error', 'Source Spreadsheet ID not configured in Settings.', ui.ButtonSet.OK);
    return;
  }
  
  const syncId = Utilities.getUuid();
  const startTime = new Date();
  
  try {
    // Open source spreadsheet
    const sourceSs = SpreadsheetApp.openById(settings.sourceSpreadsheetId);
    const sourceSheet = sourceSs.getSheetByName(settings.sourceLogSheetName);
    
    if (!sourceSheet) {
      throw new Error(`Source sheet "${settings.sourceLogSheetName}" not found`);
    }
    
    // Load source data
    const sourceData = sourceSheet.getRange(2, 1, sourceSheet.getLastRow() - 1, 24).getValues();
    
    // Load client lookup
    const clientLookup = loadClientLookup(ss);
    
    // Load existing interactions (for idempotency check)
    const existingInteractions = loadExistingInteractions(ss);
    
    // Track unmatched contacts
    const unmatchedTracker = {};
    
    // Sync interactions
    let newInteractions = 0;
    let unmatchedCount = 0;
    
    sourceData.forEach(row => {
      const messageId = row[9]; // Message ID column
      const primaryContact = row[12]; // Primary Contact Email
      const direction = row[13]; // Direction
      
      if (!messageId) return;
      
      // Match client
      const client = matchClient(primaryContact, clientLookup);
      
      if (!client && primaryContact) {
        // Track unmatched contact
        if (!unmatchedTracker[primaryContact]) {
          unmatchedTracker[primaryContact] = {
            email: primaryContact,
            domain: extractDomain(primaryContact),
            firstSeen: row[5], // Date Sent
            lastSeen: row[5],
            count: 0,
            subjects: []
          };
        }
        unmatchedTracker[primaryContact].count++;
        unmatchedTracker[primaryContact].lastSeen = row[5];
        if (row[4] && unmatchedTracker[primaryContact].subjects.length < 5) {
          unmatchedTracker[primaryContact].subjects.push(row[4].substring(0, 100));
        }
        unmatchedCount++;
      }
      
      // Create interaction ID
      const interactionId = client ? `${client.clientId}__${messageId}` : `UNMATCHED__${messageId}`;
      
      // Skip if already synced
      if (existingInteractions.has(interactionId)) return;
      
      // Append interaction
      const interactionSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
      interactionSheet.appendRow([
        interactionId,           // A - Interaction ID
        client ? client.clientId : '',  // B - Client ID
        messageId,               // C - Message ID
        row[10],                 // D - Thread ID
        row[5],                  // E - Date
        direction,               // F - Direction
        primaryContact,          // G - Primary Contact
        row[4],                  // H - Subject
        row[7],                  // I - Snippet
        row[11],                 // J - Gmail Link
        new Date().toISOString() // K - Synced At
      ]);
      
      newInteractions++;
    });
    
    // Update unmatched contacts
    updateUnmatchedContacts(ss, unmatchedTracker);
    
    // Update sync log
    logSync(ss, syncId, startTime, sourceData.length, newInteractions, unmatchedCount, 'Success', '');
    
    // Update settings
    updateSetting(ss, 'Last Sync', new Date().toISOString());
    updateSetting(ss, 'Last Sync Count', newInteractions);
    
    ui.alert('✅ Sync Complete', 
      `Synced ${newInteractions} new interactions from ${sourceData.length} source rows.\n` +
      `Unmatched contacts: ${Object.keys(unmatchedTracker).length}`,
      ui.ButtonSet.OK);
    
    // Run phase 2: rebuild thread state
    rebuildThreadState(ss);
    
    // Run phase 3: rebuild client pulse
    rebuildClientPulse(ss);
    
    // Run phase 4: generate reminders
    generateReminders(ss);
    
  } catch (e) {
    logSync(ss, syncId, startTime, 0, 0, 0, 'Failed', e.toString());
    ui.alert('❌ Sync Failed', e.toString(), ui.ButtonSet.OK);
  }
}

/**
 * Loads existing interaction IDs for idempotency check
 */
function loadExistingInteractions(ss) {
  const sheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  if (!sheet || sheet.getLastRow() < 2) return new Set();
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 1).getValues();
  return new Set(data.map(row => row[0]));
}

/**
 * Updates the UnmatchedContacts sheet
 */
function updateUnmatchedContacts(ss, unmatchedTracker) {
  const sheet = ss.getSheetByName(SHEET_NAMES.UNMATCHED_CONTACTS);
  if (!sheet) return;
  
  // Load existing unmatched
  const existing = {};
  if (sheet.getLastRow() > 1) {
    const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
    data.forEach(row => {
      existing[row[0]] = {
        firstSeen: row[2],
        count: row[4],
        subjects: (row[5] || '').split('|')
      };
    });
  }
  
  // Merge with new unmatched
  Object.keys(unmatchedTracker).forEach(email => {
    const newInfo = unmatchedTracker[email];
    
    if (existing[email]) {
      // Update existing
      existing[email].count += newInfo.count;
      existing[email].subjects = [...existing[email].subjects, ...newInfo.subjects].slice(0, 5);
    } else {
      // Add new
      sheet.appendRow([
        email,                          // A - Email
        newInfo.domain,                 // B - Domain
        newInfo.firstSeen,              // C - First Seen
        newInfo.lastSeen,               // D - Last Seen
        newInfo.count,                  // E - Interaction Count
        newInfo.subjects.join(' | '),   // F - Sample Subjects
        suggestClientName(email),       // G - Suggested Client
        'New',                          // H - Status
        new Date().toISOString()        // I - Created At
      ]);
    }
  });
}

/**
 * Suggests a client name from an email address
 */
function suggestClientName(email) {
  const domain = extractDomain(email);
  const parts = domain.split('.');
  if (parts.length >= 2) {
    return parts[0].charAt(0).toUpperCase() + parts[0].slice(1);
  }
  return domain;
}

/**
 * Extracts domain from email
 */
function extractDomain(email) {
  if (!email || !email.includes('@')) return '';
  return email.split('@')[1].toLowerCase();
}

/**
 * Logs sync result
 */
function logSync(ss, syncId, startTime, sourceRows, newInteractions, unmatchedCount, status, error) {
  const sheet = ss.getSheetByName(SHEET_NAMES.SYNC_LOG);
  if (!sheet) return;
  
  sheet.appendRow([
    syncId,
    startTime.toISOString(),
    new Date().toISOString(),
    sourceRows,
    newInteractions,
    0, // Updated Threads (will be filled by rebuild)
    0, // Updated Clients (will be filled by rebuild)
    0, // New Reminders (will be filled by generate)
    status,
    error
  ]);
}

// ============================================================================
// PHASE 2: THREAD STATE
// ============================================================================

/**
 * Rebuilds ThreadState sheet from Interactions
 */
function rebuildThreadState(ss) {
  const interactionsSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  const threadStateSheet = ss.getSheetByName(SHEET_NAMES.THREAD_STATE);
  
  if (!interactionsSheet || !threadStateSheet) return;
  
  if (interactionsSheet.getLastRow() < 2) return;
  
  // Clear existing thread state
  threadStateSheet.getRange(2, 1, threadStateSheet.getLastRow(), 13).clear();
  
  // Load all interactions
  const interactions = interactionsSheet.getRange(2, 1, interactionsSheet.getLastRow() - 1, 11).getValues();
  
  // Group by thread
  const threadGroups = {};
  interactions.forEach(row => {
    const threadId = row[3]; // Thread ID
    if (!threadId) return;
    
    if (!threadGroups[threadId]) {
      threadGroups[threadId] = {
        messages: [],
        clientId: row[1],
        subject: row[7]
      };
    }
    threadGroups[threadId].messages.push({
      messageId: row[2],
      date: new Date(row[4]),
      direction: row[5],
      contact: row[6],
      gmailLink: row[9]
    });
  });
  
  // Build thread state for each
  const settings = getSettings(ss);
  const now = new Date();
  
  Object.keys(threadGroups).forEach(threadId => {
    const group = threadGroups[threadId];
    const messages = group.messages.sort((a, b) => a.date - b.date);
    
    if (messages.length === 0) return;
    
    const firstMessage = messages[0];
    const lastMessage = messages[messages.length - 1];
    
    // Compute waiting on
    const waitingOn = lastMessage.direction === 'Outbound' ? 'Them' : 'Us';
    
    // Compute avg reply time
    let avgReplyTimeHrs = 0;
    if (messages.length > 1) {
      let totalReplyTime = 0;
      let replyCount = 0;
      
      for (let i = 1; i < messages.length; i++) {
        // Only count actual replies (direction changes)
        if (messages[i].direction !== messages[i-1].direction) {
          const diffHrs = (messages[i].date - messages[i-1].date) / (1000 * 60 * 60);
          totalReplyTime += diffHrs;
          replyCount++;
        }
      }
      
      avgReplyTimeHrs = replyCount > 0 ? Math.round(totalReplyTime / replyCount) : 0;
    }
    
    // Compute status
    const daysSinceLastMessage = Math.floor((now - lastMessage.date) / (1000 * 60 * 60 * 24));
    let status = 'Active';
    if (daysSinceLastMessage > settings.silenceCriticalDays) {
      status = 'Stale';
    } else if (waitingOn === 'Us' && daysSinceLastMessage > settings.silenceWarningDays) {
      status = 'Waiting';
    }
    
    // Compute risk score (0-100)
    let riskScore = 0;
    if (waitingOn === 'Us') {
      riskScore = Math.min(100, Math.floor(daysSinceLastMessage / settings.silenceCriticalDays * 100));
    }
    
    threadStateSheet.appendRow([
      threadId,                              // A - Thread ID
      group.clientId,                        // B - Client ID
      messages.length,                       // C - Message Count
      firstMessage.date.toISOString(),       // D - First Message Date
      lastMessage.date.toISOString(),        // E - Last Message Date
      lastMessage.contact,                   // F - Last Sender
      waitingOn,                             // G - Waiting On
      avgReplyTimeHrs,                       // H - Avg Reply Time (hrs)
      status,                                // I - Status
      riskScore,                             // J - Risk Score
      group.subject,                         // K - Subject
      lastMessage.gmailLink,                 // L - Gmail Link
      now.toISOString()                      // M - Updated At
    ]);
  });
}

// ============================================================================
// PHASE 3: CLIENT PULSE
// ============================================================================

/**
 * Rebuilds ClientPulse sheet from Interactions and ThreadState
 */
function rebuildClientPulse(ss) {
  const clientsSheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  const interactionsSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  const threadStateSheet = ss.getSheetByName(SHEET_NAMES.THREAD_STATE);
  const clientPulseSheet = ss.getSheetByName(SHEET_NAMES.CLIENT_PULSE);
  
  if (!clientsSheet || !interactionsSheet || !clientPulseSheet) return;
  
  // Clear existing pulse
  if (clientPulseSheet.getLastRow() > 1) {
    clientPulseSheet.getRange(2, 1, clientPulseSheet.getLastRow(), 17).clear();
  }
  
  // Load clients
  const clientsData = clientsSheet.getLastRow() > 1 
    ? clientsSheet.getRange(2, 1, clientsSheet.getLastRow() - 1, 13).getValues()
    : [];
  
  // Load interactions
  const interactionsData = interactionsSheet.getLastRow() > 1
    ? interactionsSheet.getRange(2, 1, interactionsSheet.getLastRow() - 1, 11).getValues()
    : [];
  
  // Load thread state
  const threadStateData = threadStateSheet && threadStateSheet.getLastRow() > 1
    ? threadStateSheet.getRange(2, 1, threadStateSheet.getLastRow() - 1, 13).getValues()
    : [];
  
  const settings = getSettings(ss);
  const now = new Date();
  
  // Group interactions by client
  const clientInteractions = {};
  interactionsData.forEach(row => {
    const clientId = row[1]; // Client ID
    if (!clientId) return;
    
    if (!clientInteractions[clientId]) {
      clientInteractions[clientId] = {
        all: [],
        inbound: [],
        outbound: []
      };
    }
    
    const interaction = {
      date: new Date(row[4]),
      direction: row[5],
      threadId: row[3]
    };
    
    clientInteractions[clientId].all.push(interaction);
    if (interaction.direction === 'Inbound') {
      clientInteractions[clientId].inbound.push(interaction);
    } else if (interaction.direction === 'Outbound') {
      clientInteractions[clientId].outbound.push(interaction);
    }
  });
  
  // Group threads by client
  const clientThreads = {};
  threadStateData.forEach(row => {
    const clientId = row[1]; // Client ID
    if (!clientId) return;
    
    if (!clientThreads[clientId]) {
      clientThreads[clientId] = [];
    }
    
    clientThreads[clientId].push({
      threadId: row[0],
      status: row[8],
      waitingOn: row[6],
      riskScore: row[9],
      lastDate: new Date(row[4])
    });
  });
  
  // Build pulse for each client
  clientsData.forEach(clientRow => {
    const clientId = clientRow[0];
    const clientName = clientRow[1];
    const priority = clientRow[6] || 'Medium';
    
    const interactions = clientInteractions[clientId] || { all: [], inbound: [], outbound: [] };
    const threads = clientThreads[clientId] || [];
    
    // Sort by date descending
    interactions.all.sort((a, b) => b.date - a.date);
    interactions.inbound.sort((a, b) => b.date - a.date);
    interactions.outbound.sort((a, b) => b.date - a.date);
    
    // Compute metrics
    const lastInteraction = interactions.all.length > 0 ? interactions.all[0].date : null;
    const lastInbound = interactions.inbound.length > 0 ? interactions.inbound[0].date : null;
    const lastOutbound = interactions.outbound.length > 0 ? interactions.outbound[0].date : null;
    
    // Compute silence days
    let currentSilenceDays = 0;
    if (lastInteraction) {
      currentSilenceDays = Math.floor((now - lastInteraction) / (1000 * 60 * 60 * 24));
    }
    
    // Determine typical silence (could be computed from historical data, using settings for now)
    const typicalSilenceDays = settings.typicalSilenceDays;
    
    // Compute drift score
    let silenceDriftScore = 0;
    if (typicalSilenceDays > 0) {
      silenceDriftScore = (currentSilenceDays - typicalSilenceDays) / typicalSilenceDays;
    }
    
    // Adjust for priority
    let priorityMultiplier = 1;
    if (priority === 'High') priorityMultiplier = 1.5;
    else if (priority === 'Low') priorityMultiplier = 0.75;
    
    const adjustedDriftScore = silenceDriftScore * priorityMultiplier;
    
    // Determine temperature
    let temperature = 'Warm';
    if (currentSilenceDays <= 7) temperature = 'Hot';
    else if (currentSilenceDays <= 14) temperature = 'Warm';
    else if (currentSilenceDays <= 21) temperature = 'Cool';
    else temperature = 'Cold';
    
    // Determine status
    let status = 'Healthy';
    if (currentSilenceDays > settings.silenceCriticalDays) status = 'Overdue';
    else if (currentSilenceDays > settings.silenceWarningDays) status = 'Critical';
    else if (currentSilenceDays > settings.typicalSilenceDays) status = 'Warning';
    
    // Waiting on
    const activeThread = threads.find(t => t.status === 'Active' || t.status === 'Waiting');
    const waitingOn = activeThread ? activeThread.waitingOn : 'N/A';
    
    // Active threads
    const activeThreads = threads.filter(t => t.status === 'Active').length;
    
    // Total interactions
    const totalInteractions = interactions.all.length;
    
    // Reminder score
    let reminderScore = Math.min(100, Math.floor(currentSilenceDays * 3));
    if (waitingOn === 'Us') reminderScore = Math.min(100, reminderScore + 20);
    if (priority === 'High') reminderScore = Math.min(100, reminderScore + 10);
    
    // Recommended action
    let recommendedAction = 'No action needed';
    if (status === 'Overdue') recommendedAction = 'Urgent follow-up required';
    else if (status === 'Critical') recommendedAction = 'Follow up within 24 hours';
    else if (status === 'Warning') recommendedAction = 'Consider reaching out';
    else if (waitingOn === 'Them' && currentSilenceDays > 7) recommendedAction = 'Waiting for reply';
    
    // Update clients sheet with computed values
    clientRow[7] = lastInteraction ? lastInteraction.toISOString() : ''; // First Contact
    clientRow[8] = lastInteraction ? lastInteraction.toISOString() : ''; // Last Contact
    clientRow[9] = totalInteractions; // Total Interactions
    
    clientPulseSheet.appendRow([
      clientId,
      clientName,
      lastInteraction ? lastInteraction.toISOString() : '',
      lastInbound ? lastInbound.toISOString() : '',
      lastOutbound ? lastOutbound.toISOString() : '',
      typicalSilenceDays,
      currentSilenceDays,
      Math.round(silenceDriftScore * 100) / 100,
      Math.round(adjustedDriftScore * 100) / 100,
      temperature,
      status,
      waitingOn,
      activeThreads,
      totalInteractions,
      reminderScore,
      recommendedAction,
      now.toISOString()
    ]);
  });
  
  // Update clients sheet
  if (clientsData.length > 0) {
    clientsSheet.getRange(2, 1, clientsData.length, clientsData[0].length).setValues(clientsData);
  }
}

// ============================================================================
// PHASE 4: REMINDERS
// ============================================================================

/**
 * Generates reminders from ClientPulse
 */
function generateReminders(ss) {
  const clientPulseSheet = ss.getSheetByName(SHEET_NAMES.CLIENT_PULSE);
  const remindersSheet = ss.getSheetByName(SHEET_NAMES.REMINDERS);
  
  if (!clientPulseSheet || !remindersSheet) return;
  
  if (clientPulseSheet.getLastRow() < 2) return;
  
  const settings = getSettings(ss);
  const now = new Date();
  const pulseData = clientPulseSheet.getRange(2, 1, clientPulseSheet.getLastRow() - 1, 17).getValues();
  
  // Load existing pending reminders
  const existingReminders = new Set();
  if (remindersSheet.getLastRow() > 1) {
    const reminderData = remindersSheet.getRange(2, 1, remindersSheet.getLastRow() - 1, 12).getValues();
    reminderData.forEach(row => {
      if (row[9] === 'Pending') { // Status = Pending
        existingReminders.add(`${row[1]}__${row[3]}`); // ClientID__Type
      }
    });
  }
  
  let newReminders = 0;
  
  pulseData.forEach(row => {
    const clientId = row[0];
    const clientName = row[1];
    const currentSilenceDays = row[6] || 0;
    const status = row[10];
    const waitingOn = row[11];
    const reminderScore = row[14] || 0;
    
    // Determine reminder type and priority
    let reminderType = null;
    let priority = 'Medium';
    let message = '';
    
    if (status === 'Overdue') {
      reminderType = 'Critical';
      priority = 'High';
      message = `${clientName}: Overdue for follow-up (${currentSilenceDays} days since last contact)`;
    } else if (status === 'Critical') {
      reminderType = 'Stale';
      priority = 'High';
      message = `${clientName}: Critical - ${currentSilenceDays} days since last contact`;
    } else if (status === 'Warning' && waitingOn === 'Us') {
      reminderType = 'FollowUp';
      priority = 'Medium';
      message = `${clientName}: Follow-up suggested (${currentSilenceDays} days)`;
    } else if (waitingOn === 'Them' && currentSilenceDays > 7) {
      reminderType = 'Waiting';
      priority = 'Low';
      message = `${clientName}: Waiting for reply (${currentSilenceDays} days)`;
    }
    
    if (reminderType && !existingReminders.has(`${clientId}__${reminderType}`)) {
      remindersSheet.appendRow([
        Utilities.getUuid(),           // A - Reminder ID
        clientId,                       // B - Client ID
        clientName,                     // C - Client Name
        reminderType,                   // D - Type
        priority,                       // E - Priority
        message,                        // F - Message
        currentSilenceDays,             // G - Days Since Contact
        '',                             // H - Thread ID
        '',                             // I - Gmail Link
        'Pending',                      // J - Status
        now.toISOString(),              // K - Created At
        ''                              // L - Completed At
      ]);
      newReminders++;
    }
  });
  
  return newReminders;
}

/**
 * Shows current status
 */
function showStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settings = getSettings(ss);
  
  const interactionsSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  const remindersSheet = ss.getSheetByName(SHEET_NAMES.REMINDERS);
  const clientsSheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  
  const interactionCount = interactionsSheet ? Math.max(0, interactionsSheet.getLastRow() - 1) : 0;
  const clientCount = clientsSheet ? Math.max(0, clientsSheet.getLastRow() - 1) : 0;
  const reminderCount = remindersSheet ? Math.max(0, remindersSheet.getLastRow() - 1) : 0;
  
  SpreadsheetApp.getUi().alert(
    '📋 Client Reminder Status',
    `Last Sync: ${settings.lastSync || 'Never'}\n` +
    `Interactions: ${interactionCount}\n` +
    `Clients: ${clientCount}\n` +
    `Pending Reminders: ${reminderCount}`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

/**
 * Shows unmatched contacts count
 */
function showUnmatchedCount() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAMES.UNMATCHED_CONTACTS);
  
  if (!sheet || sheet.getLastRow() < 2) {
    SpreadsheetApp.getUi().alert('Unmatched Contacts', 'No unmatched contacts found.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  const data = sheet.getRange(2, 1, sheet.getLastRow() - 1, 9).getValues();
  const newCount = data.filter(row => row[7] === 'New').length;
  const totalCount = data.length;
  
  SpreadsheetApp.getUi().alert(
    'Unmatched Contacts',
    `Total: ${totalCount}\nNew: ${newCount}\n\nCheck the UnmatchedContacts sheet to add them as clients.`,
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}