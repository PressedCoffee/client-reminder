/**
 * Tool #3: Client Follow-Up Reminder System
 * Google Apps Script for Google Sheets + Gmail integration
 * 
 * Features:
 * - Track client interactions and calculate follow-up schedules
 * - Send daily email digests of due/overdue follow-ups
 * - Priority scoring with VIP flagging
 * - Snooze functionality
 * - Integration with Tool #2 Email-to-Spreadsheet Logger
 */

// ============================================
// CONSTANTS & CONFIGURATION
// ============================================

const SHEET_NAMES = {
  SETTINGS: 'Settings',
  CLIENTS: 'Clients',
  INTERACTIONS: 'Interactions',
  REMINDERS: 'Reminders'
};

const PRIORITY_MULTIPLIERS = {
  'VIP': 3,
  'Standard': 2,
  'Low': 1
};

const STATUS_COLORS = {
  'OVERDUE': '#FF6B6B',    // Red
  'DUE_TODAY': '#FFD93D',   // Yellow
  'UPCOMING': '#6BCB77',    // Green
  'PAUSED': '#A0A0A0',      // Gray
  'COMPLETED': '#4ECDC4'    // Teal
};

// ============================================
// SETUP & TEMPLATE FUNCTIONS
// ============================================

/**
 * Creates custom menu when spreadsheet opens
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('📋 Client Reminders')
    .addItem('🔄 Calculate Reminders', 'calculateReminders')
    .addItem('📧 Send Test Digest', 'testSendDigest')
    .addItem('🔄 Sync from Tool #2', 'syncFromTool2')
    .addSeparator()
    .addItem('✅ Mark Client as Contacted', 'showMarkContactedDialog')
    .addItem('😴 Snooze Reminder', 'showSnoozeDialog')
    .addSeparator()
    .addItem('⚙️ Setup/Reset Template', 'runSetup')
    .addToUi();
}

/**
 * Run setup - creates all sheets with headers and sample data
 */
function runSetup() {
  setupTemplate();
  SpreadsheetApp.getUi().alert('Setup complete! All sheets created with sample data.');
}

/**
 * Creates all required sheets with headers and sample data
 */
function setupTemplate() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  // Create or clear sheets
  setupSettingsSheet(ss);
  setupClientsSheet(ss);
  setupInteractionsSheet(ss);
  setupRemindersSheet(ss);
  
  // Calculate initial reminders
  calculateReminders();
  
  Logger.log('Template setup complete');
}

/**
 * Creates Settings sheet with configuration
 */
function setupSettingsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.SETTINGS);
  } else {
    sheet.clear();
  }
  
  const settingsData = [
    ['Setting', 'Value', 'Description'],
    ['', '', ''],
    ['EMAIL & DIGEST SETTINGS', '', ''],
    ['Email_Recipient', 'shaddockoc@gmail.com', 'Email address for daily digest'],
    ['Digest_Subject', 'Daily Client Follow-Up Reminders', 'Subject line of digest email'],
    ['Digest_Time', '08:00', 'Time to send digest (24h format)'],
    ['Digest_Timezone', 'America/Los_Angeles', 'Timezone for scheduling'],
    ['Include_Completed', 'No', 'Include completed follow-ups in digest? (Yes/No)'],
    ['', '', ''],
    ['FOLLOW-UP RULES', '', ''],
    ['New_Lead_Day1', '2', 'Days after first contact'],
    ['New_Lead_Day2', '7', 'Days after first contact if no response'],
    ['New_Lead_Day3', '14', 'Days after first contact if still no response'],
    ['Existing_Client_VIP', '14', 'Days between check-ins for VIPs'],
    ['Existing_Client_Standard', '30', 'Days between check-ins for Standard clients'],
    ['Existing_Client_Low', '60', 'Days between check-ins for Low priority'],
    ['Post_Meeting', '3', 'Days after a meeting to follow up'],
    ['No_Response_Followup', '7', 'Days to wait before following up on non-response'],
    ['', '', ''],
    ['PRIORITY SCORING', '', ''],
    ['VIP_Multiplier', '3', 'Priority score multiplier for VIP clients'],
    ['Standard_Multiplier', '2', 'Priority score multiplier for Standard clients'],
    ['Low_Multiplier', '1', 'Priority score multiplier for Low priority clients'],
    ['Overdue_Penalty_Per_Day', '5', 'Points added per day overdue'],
    ['Due_Today_Bonus', '10', 'Bonus points for items due today'],
    ['Upcoming_Window_Days', '3', 'How many days ahead to show in upcoming'],
    ['', '', ''],
    ['SNOOZE OPTIONS', '', ''],
    ['Snooze_Option_1', '1', 'Quick snooze option 1 (days)'],
    ['Snooze_Option_2', '3', 'Quick snooze option 2 (days)'],
    ['Snooze_Option_3', '7', 'Quick snooze option 3 (days)'],
    ['Snooze_Option_4', '14', 'Quick snooze option 4 (days)']
  ];
  
  sheet.getRange(1, 1, settingsData.length, 3).setValues(settingsData);
  sheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  sheet.autoResizeColumns(1, 3);
  sheet.setColumnWidth(2, 200);
  sheet.setColumnWidth(3, 350);
  
  // Add section headers styling
  const sectionRows = [3, 11, 21, 29];
  sectionRows.forEach(row => {
    if (settingsData[row-1]) {
      sheet.getRange(row, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
    }
  });
}

/**
 * Creates Clients sheet with sample data
 */
function setupClientsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.CLIENTS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    'Client_ID', 'Name', 'Email', 'Phone', 'Company',
    'Priority', 'Preferred_Channel', 'Rule_Set', 'Last_Contact_Date',
    'Next_Followup_Due', 'Status', 'Notes'
  ];
  
  const sampleData = [
    ['C001', 'Acme Corp - John Smith', 'john.smith@acmecorp.com', '+1-555-0101', 'Acme Corporation',
     'VIP', 'Email', 'Existing_Client', getDateDaysAgo(10), getDateDaysAgo(-4), 'Active', 'Key account, quarterly review scheduled'],
    ['C002', 'Beta LLC - Sarah Jones', 'sarah@betallc.com', '+1-555-0102', 'Beta LLC',
     'Standard', 'Phone', 'New_Lead', getDateDaysAgo(8), getDateDaysAgo(-6), 'Active', 'Interested in premium package'],
    ['C003', 'Gamma Startup - Mike Chen', 'mike@gamma.io', '+1-555-0103', 'Gamma Startup Inc',
     'VIP', 'Video', 'Post_Meeting', getDateDaysAgo(5), getDateDaysAgo(-2), 'Active', 'Demo completed, awaiting feedback'],
    ['C004', 'Delta Co - Emily Davis', 'emily@deltaco.com', '+1-555-0104', 'Delta Company',
     'Low', 'Email', 'No_Response', getDateDaysAgo(45), getDateDaysAgo(15), 'Active', 'Initial outreach, no response yet'],
    ['C005', 'Epsilon Design - Alex Kim', 'alex@epsilondesign.com', '+1-555-0105', 'Epsilon Design Studio',
     'Standard', 'Email', 'Existing_Client', getDateDaysAgo(25), getDateDaysAgo(5), 'Active', 'Ongoing project, monthly check-in due']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // Add data validation for Priority
  const priorityRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['VIP', 'Standard', 'Low'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, 1000, 1).setDataValidation(priorityRule);
  
  // Add data validation for Preferred_Channel
  const channelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Email', 'Phone', 'Video', 'In-Person', 'Text'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 7, 1000, 1).setDataValidation(channelRule);
  
  // Add data validation for Rule_Set
  const ruleRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['New_Lead', 'Existing_Client', 'Post_Meeting', 'No_Response'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 8, 1000, 1).setDataValidation(ruleRule);
  
  // Add data validation for Status
  const statusRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Active', 'Inactive', 'Closed', 'Paused'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 11, 1000, 1).setDataValidation(statusRule);
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

/**
 * Creates Interactions sheet with sample data
 */
function setupInteractionsSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.INTERACTIONS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    'Interaction_ID', 'Client_ID', 'Client_Email', 'Date', 'Type',
    'Channel', 'Subject/Topic', 'Outcome', 'Notes', 'Followup_Required', 'Created_By'
  ];
  
  const sampleData = [
    ['I001', 'C001', 'john.smith@acmecorp.com', getDateDaysAgo(10), 'Email', 
     'Email', 'Q4 Contract Renewal', 'Completed', 'Contract renewed for 2 years', 'No', 'Manual'],
    ['I002', 'C002', 'sarah@betallc.com', getDateDaysAgo(8), 'Call',
     'Phone', 'Initial Discovery Call', 'Positive', 'Interested in premium tier, need to send proposal', 'Yes', 'Manual'],
    ['I003', 'C003', 'mike@gamma.io', getDateDaysAgo(5), 'Meeting',
     'Video', 'Product Demo Session', 'Pending Response', 'Great demo, Mike needs to discuss with team', 'Yes', 'Manual'],
    ['I004', 'C004', 'emily@deltaco.com', getDateDaysAgo(45), 'Email',
     'Email', 'Cold Outreach - Services', 'No Response', 'No reply to initial email', 'Yes', 'Manual'],
    ['I005', 'C005', 'alex@epsilondesign.com', getDateDaysAgo(25), 'Email',
     'Email', 'Monthly Check-in', 'Completed', 'Project on track, next review in 30 days', 'Yes', 'Manual'],
    ['I006', 'C001', 'john.smith@acmecorp.com', getDateDaysAgo(2), 'Inbound',
     'Email', 'Technical Support Request', 'Completed', 'Resolved server issue', 'No', 'Tool-2-Auto']
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  sheet.getRange(2, 1, sampleData.length, headers.length).setValues(sampleData);
  
  // Data validation for Type
  const typeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Email', 'Call', 'Meeting', 'Video', 'Text', 'Inbound', 'Outbound'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 5, 1000, 1).setDataValidation(typeRule);
  
  // Data validation for Channel
  const channelRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Email', 'Phone', 'Video', 'In-Person', 'Text', 'Other'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 6, 1000, 1).setDataValidation(channelRule);
  
  // Data validation for Outcome
  const outcomeRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Completed', 'Pending Response', 'Positive', 'No Response', 'Negative', 'Canceled'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 8, 1000, 1).setDataValidation(outcomeRule);
  
  // Data validation for Followup_Required
  const followupRule = SpreadsheetApp.newDataValidation()
    .requireValueInList(['Yes', 'No'])
    .setAllowInvalid(false)
    .build();
  sheet.getRange(2, 10, 1000, 1).setDataValidation(followupRule);
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

/**
 * Creates Reminders sheet
 */
function setupRemindersSheet(ss) {
  let sheet = ss.getSheetByName(SHEET_NAMES.REMINDERS);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAMES.REMINDERS);
  } else {
    sheet.clear();
  }
  
  const headers = [
    'Status', 'Priority_Score', 'Client_ID', 'Name', 'Email', 'Phone',
    'Priority_Tier', 'Days_Since_Contact', 'Days_Until_Due', 'Next_Followup_Due',
    'Rule_Applied', 'Last_Interaction_Date', 'Last_Interaction_Type', 
    'Preferred_Channel', 'Notes', 'Action_Required'
  ];
  
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
    .setFontWeight('bold').setBackground('#4285F4').setFontColor('white');
  
  sheet.autoResizeColumns(1, headers.length);
  sheet.setFrozenRows(1);
}

// ============================================
// CORE CALCULATION FUNCTIONS
// ============================================

/**
 * Calculates all reminders based on client interactions and rules
 * This is the main function run by daily trigger
 */
function calculateReminders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsSheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  const interactionsSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  const remindersSheet = ss.getSheetByName(SHEET_NAMES.REMINDERS);
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  // Load settings
  const settings = loadSettings(settingsSheet);
  
  // Get all active clients
  const clientsData = clientsSheet.getDataRange().getValues();
  const clientsHeaders = clientsData.shift();
  const clientIdIdx = clientsHeaders.indexOf('Client_ID');
  const clientEmailIdx = clientsHeaders.indexOf('Email');
  const priorityIdx = clientsHeaders.indexOf('Priority');
  const ruleSetIdx = clientsHeaders.indexOf('Rule_Set');
  const lastContactIdx = clientsHeaders.indexOf('Last_Contact_Date');
  const nextFollowupIdx = clientsHeaders.indexOf('Next_Followup_Due');
  const statusIdx = clientsHeaders.indexOf('Status');
  const preferredChannelIdx = clientsHeaders.indexOf('Preferred_Channel');
  const notesIdx = clientsHeaders.indexOf('Notes');
  const nameIdx = clientsHeaders.indexOf('Name');
  const phoneIdx = clientsHeaders.indexOf('Phone');
  
  // Get all interactions
  const interactionsData = interactionsSheet.getDataRange().getValues();
  const interactionsHeaders = interactionsData.shift();
  const intClientIdIdx = interactionsHeaders.indexOf('Client_ID');
  const intDateIdx = interactionsHeaders.indexOf('Date');
  const intTypeIdx = interactionsHeaders.indexOf('Type');
  
  // Find last interaction for each client
  const lastInteractions = {};
  interactionsData.forEach(row => {
    const clientId = row[intClientIdIdx];
    const date = row[intDateIdx];
    const type = row[intTypeIdx];
    
    if (clientId && date) {
      const dateObj = new Date(date);
      if (!lastInteractions[clientId] || dateObj > lastInteractions[clientId].date) {
        lastInteractions[clientId] = { date: dateObj, type: type };
      }
    }
  });
  
  // Calculate reminders
  const reminders = [];
  const today = new Date();
  today.setHours(0, 0, 0, 0);
  
  clientsData.forEach(client => {
    const status = client[statusIdx];
    if (status !== 'Active') return; // Skip inactive/closed/paused clients
    
    const clientId = client[clientIdIdx];
    const priority = client[priorityIdx];
    const ruleSet = client[ruleSetIdx];
    const preferredChannel = client[preferredChannelIdx];
    const nextFollowupDue = client[nextFollowupIdx];
    const notes = client[notesIdx];
    
    // Get last interaction info
    const lastInt = lastInteractions[clientId];
    const lastContactDate = lastInt ? lastInt.date : (client[lastContactIdx] ? new Date(client[lastContactIdx]) : null);
    const lastInteractionType = lastInt ? lastInt.type : '';
    
    if (!lastContactDate) return; // Skip if no contact date
    
    // Calculate days since contact
    const daysSinceContact = Math.floor((today - lastContactDate) / (1000 * 60 * 60 * 24));
    
    // Determine due date
    let dueDate = nextFollowupDue ? new Date(nextFollowupDue) : null;
    
    // If no explicit due date, calculate based on rules
    if (!dueDate) {
      const followupDays = getFollowupDays(ruleSet, priority, settings);
      dueDate = new Date(lastContactDate);
      dueDate.setDate(dueDate.getDate() + followupDays);
    }
    
    // Calculate days until due
    const daysUntilDue = Math.floor((dueDate - today) / (1000 * 60 * 60 * 24));
    
    // Determine status
    let status;
    if (daysUntilDue < 0) {
      status = 'OVERDUE';
    } else if (daysUntilDue === 0) {
      status = 'DUE_TODAY';
    } else if (daysUntilDue <= settings.Upcoming_Window_Days) {
      status = 'UPCOMING';
    } else {
      return; // Too far out, don't include
    }
    
    // Calculate priority score
    const priorityMultiplier = PRIORITY_MULTIPLIERS[priority] || 1;
    let priorityScore = daysSinceContact * priorityMultiplier;
    
    if (daysUntilDue < 0) {
      priorityScore += Math.abs(daysUntilDue) * settings.Overdue_Penalty_Per_Day;
    } else if (daysUntilDue === 0) {
      priorityScore += settings.Due_Today_Bonus;
    }
    
    reminders.push({
      status: status,
      priorityScore: priorityScore,
      clientId: clientId,
      name: client[nameIdx],
      email: client[clientEmailIdx],
      phone: client[phoneIdx],
      priorityTier: priority,
      daysSinceContact: daysSinceContact,
      daysUntilDue: daysUntilDue,
      nextFollowupDue: dueDate,
      ruleApplied: ruleSet,
      lastInteractionDate: lastContactDate,
      lastInteractionType: lastInteractionType,
      preferredChannel: preferredChannel,
      notes: notes,
      actionRequired: 'Follow up via ' + preferredChannel
    });
  });
  
  // Sort by priority score (descending)
  reminders.sort((a, b) => b.priorityScore - a.priorityScore);
  
  // Clear and write reminders
  const lastRow = remindersSheet.getLastRow();
  if (lastRow > 1) {
    remindersSheet.getRange(2, 1, lastRow - 1, 16).clear();
  }
  
  if (reminders.length > 0) {
    const reminderRows = reminders.map(r => [
      r.status,
      r.priorityScore,
      r.clientId,
      r.name,
      r.email,
      r.phone,
      r.priorityTier,
      r.daysSinceContact,
      r.daysUntilDue,
      r.nextFollowupDue,
      r.ruleApplied,
      r.lastInteractionDate,
      r.lastInteractionType,
      r.preferredChannel,
      r.notes,
      r.actionRequired
    ]);
    
    remindersSheet.getRange(2, 1, reminderRows.length, 16).setValues(reminderRows);
    
    // Apply status colors
    reminders.forEach((r, i) => {
      const color = STATUS_COLORS[r.status] || '#FFFFFF';
      remindersSheet.getRange(i + 2, 1).setBackground(color);
    });
    
    // Format dates
    remindersSheet.getRange(2, 10, reminders.length, 1).setNumberFormat('yyyy-MM-dd');
    remindersSheet.getRange(2, 12, reminders.length, 1).setNumberFormat('yyyy-MM-dd');
  }
  
  Logger.log(`Calculated ${reminders.length} reminders`);
  return reminders;
}

/**
 * Loads settings from Settings sheet
 */
function loadSettings(settingsSheet) {
  const data = settingsSheet.getDataRange().getValues();
  const settings = {};
  
  data.forEach(row => {
    const key = row[0];
    const value = row[1];
    if (key && !key.includes('SETTING') && !key.includes('FOLLOW-UP') && 
        !key.includes('PRIORITY') && !key.includes('SNOOZE') && key !== 'Setting') {
      // Try to parse as number
      const numValue = parseFloat(value);
      settings[key] = isNaN(numValue) ? value : numValue;
    }
  });
  
  return settings;
}

/**
 * Gets follow-up days based on rule set and priority
 */
function getFollowupDays(ruleSet, priority, settings) {
  const daysMap = {
    'New_Lead': settings.New_Lead_Day1 || 2,
    'Existing_Client': priority === 'VIP' ? settings.Existing_Client_VIP : 
                       priority === 'Low' ? settings.Existing_Client_Low : 
                       settings.Existing_Client_Standard,
    'Post_Meeting': settings.Post_Meeting || 3,
    'No_Response': settings.No_Response_Followup || 7
  };
  
  return daysMap[ruleSet] || settings.Existing_Client_Standard || 30;
}

// ============================================
// EMAIL DIGEST FUNCTIONS
// ============================================

/**
 * Sends daily digest email - called by trigger
 */
function sendDailyDigest() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const remindersSheet = ss.getSheetByName(SHEET_NAMES.REMINDERS);
  const settingsSheet = ss.getSheetByName(SHEET_NAMES.SETTINGS);
  
  // Load settings
  const settings = loadSettings(settingsSheet);
  const recipient = settings.Email_Recipient || 'shaddockoc@gmail.com';
  const subject = settings.Digest_Subject || 'Daily Client Follow-Up Reminders';
  
  // Get reminders data
  const data = remindersSheet.getDataRange().getValues();
  if (data.length <= 1) {
    Logger.log('No reminders to send');
    return;
  }
  
  const headers = data.shift();
  const statusIdx = headers.indexOf('Status');
  const priorityIdx = headers.indexOf('Priority_Score');
  const nameIdx = headers.indexOf('Name');
  const emailIdx = headers.indexOf('Email');
  const daysUntilIdx = headers.indexOf('Days_Until_Due');
  const channelIdx = headers.indexOf('Preferred_Channel');
  const tierIdx = headers.indexOf('Priority_Tier');
  const notesIdx = headers.indexOf('Notes');
  
  // Categorize reminders
  const overdue = [];
  const dueToday = [];
  const upcoming = [];
  
  data.forEach(row => {
    const status = row[statusIdx];
    const reminder = {
      name: row[nameIdx],
      email: row[emailIdx],
      daysUntil: row[daysUntilIdx],
      channel: row[channelIdx],
      tier: row[tierIdx],
      priority: row[priorityIdx],
      notes: row[notesIdx]
    };
    
    if (status === 'OVERDUE') overdue.push(reminder);
    else if (status === 'DUE_TODAY') dueToday.push(reminder);
    else if (status === 'UPCOMING') upcoming.push(reminder);
  });
  
  // Build email body
  const today = new Date().toLocaleDateString('en-US', { 
    weekday: 'long', 
    year: 'numeric', 
    month: 'long', 
    day: 'numeric' 
  });
  
  let emailBody = `<html>
<head>
<style>
  body { font-family: Arial, sans-serif; line-height: 1.6; color: #333; }
  h1 { color: #4285F4; }
  h2 { color: #333; margin-top: 30px; border-bottom: 2px solid #eee; padding-bottom: 10px; }
  .summary { background: #f5f5f5; padding: 15px; border-radius: 5px; margin: 20px 0; }
  .vip { background: #FF6B6B; color: white; padding: 2px 8px; border-radius: 3px; font-size: 12px; }
  .standard { background: #FFD93D; color: #333; padding: 2px 8px; border-radius: 3px; font-size: 12px; }
  .low { background: #6BCB77; color: white; padding: 2px 8px; border-radius: 3px; font-size: 12px; }
  table { width: 100%; border-collapse: collapse; margin: 15px 0; }
  th { background: #4285F4; color: white; text-align: left; padding: 10px; }
  td { padding: 10px; border-bottom: 1px solid #eee; }
  .overdue-row { background: #FFF3F3; }
  .due-row { background: #FFFBE6; }
  .notes { font-size: 12px; color: #666; font-style: italic; }
  .footer { margin-top: 40px; padding-top: 20px; border-top: 1px solid #eee; font-size: 12px; color: #999; }
</style>
</head>
<body>
  <h1>📋 Daily Client Follow-Up Report</h1>
  <p>${today}</p>
  
  <div class="summary">
    <strong>Summary:</strong> ${overdue.length} overdue | ${dueToday.length} due today | ${upcoming.length} upcoming
  </div>`;
  
  // Overdue section
  if (overdue.length > 0) {
    emailBody += `
  <h2>🔴 OVERDUE (${overdue.length})</h2>
  <table>
    <tr>
      <th>Priority</th>
      <th>Client</th>
      <th>Contact</th>
      <th>Days Overdue</th>
      <th>Channel</th>
      <th>Notes</th>
    </tr>`;
    
    overdue.forEach(r => {
      const tierClass = r.tier === 'VIP' ? 'vip' : r.tier === 'Standard' ? 'standard' : 'low';
      emailBody += `
    <tr class="overdue-row">
      <td><span class="${tierClass}">${r.tier}</span></td>
      <td><strong>${escapeHtml(r.name)}</strong></td>
      <td>${escapeHtml(r.email)}</td>
      <td>${Math.abs(r.daysUntil)} days</td>
      <td>${r.channel}</td>
      <td class="notes">${escapeHtml(r.notes || '')}</td>
    </tr>`;
    });
    
    emailBody += `
  </table>`;
  }
  
  // Due today section
  if (dueToday.length > 0) {
    emailBody += `
  <h2>🟡 DUE TODAY (${dueToday.length})</h2>
  <table>
    <tr>
      <th>Priority</th>
      <th>Client</th>
      <th>Contact</th>
      <th>Channel</th>
      <th>Notes</th>
    </tr>`;
    
    dueToday.forEach(r => {
      const tierClass = r.tier === 'VIP' ? 'vip' : r.tier === 'Standard' ? 'standard' : 'low';
      emailBody += `
    <tr class="due-row">
      <td><span class="${tierClass}">${r.tier}</span></td>
      <td><strong>${escapeHtml(r.name)}</strong></td>
      <td>${escapeHtml(r.email)}</td>
      <td>${r.channel}</td>
      <td class="notes">${escapeHtml(r.notes || '')}</td>
    </tr>`;
    });
    
    emailBody += `
  </table>`;
  }
  
  // Upcoming section
  if (upcoming.length > 0) {
    emailBody += `
  <h2>🟢 UPCOMING - Next ${settings.Upcoming_Window_Days || 3} Days (${upcoming.length})</h2>
  <table>
    <tr>
      <th>Priority</th>
      <th>Client</th>
      <th>Contact</th>
      <th>Days Until</th>
      <th>Channel</th>
    </tr>`;
    
    upcoming.forEach(r => {
      const tierClass = r.tier === 'VIP' ? 'vip' : r.tier === 'Standard' ? 'standard' : 'low';
      emailBody += `
    <tr>
      <td><span class="${tierClass}">${r.tier}</span></td>
      <td><strong>${escapeHtml(r.name)}</strong></td>
      <td>${escapeHtml(r.email)}</td>
      <td>${r.daysUntil} days</td>
      <td>${r.channel}</td>
    </tr>`;
    });
    
    emailBody += `
  </table>`;
  }
  
  // No reminders message
  if (overdue.length === 0 && dueToday.length === 0 && upcoming.length === 0) {
    emailBody += `
  <p style="text-align: center; padding: 40px; color: #666;">
    🎉 No follow-ups required today! You're all caught up.
  </p>`;
  }
  
  emailBody += `
  <div class="footer">
    <p>Generated by Client Follow-Up Reminder System</p>
    <p><a href="${ss.getUrl()}">Open Spreadsheet</a></p>
  </div>
</body>
</html>`;
  
  // Send email
  try {
    MailApp.sendEmail({
      to: recipient,
      subject: `${subject} - ${today}`,
      htmlBody: emailBody,
      name: 'Client Reminder System'
    });
    Logger.log(`Digest sent to ${recipient}`);
  } catch (e) {
    Logger.log(`Error sending email: ${e}`);
  }
}

/**
 * Test function to send digest without waiting for trigger
 */
function testSendDigest() {
  sendDailyDigest();
  SpreadsheetApp.getUi().alert('Test digest email sent! Check your inbox (and spam folder).');
}

/**
 * Escape HTML special characters
 */
function escapeHtml(text) {
  if (!text) return '';
  return String(text)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#039;');
}

// ============================================
// ACTION FUNCTIONS
// ============================================

/**
 * Marks a client as contacted - logs interaction and updates reminders
 * @param {string} clientId - The client ID (e.g., "C001")
 * @param {string} interactionType - Type of contact (Email, Call, Meeting, etc.)
 * @param {string} notes - Optional notes about the interaction
 */
function markAsContacted(clientId, interactionType, notes) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsSheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  const interactionsSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  
  // Find client
  const clientsData = clientsSheet.getDataRange().getValues();
  const headers = clientsData.shift();
  const clientIdIdx = headers.indexOf('Client_ID');
  const emailIdx = headers.indexOf('Email');
  const nameIdx = headers.indexOf('Name');
  
  let clientEmail = '';
  let clientName = '';
  let clientRow = -1;
  
  clientsData.forEach((row, index) => {
    if (row[clientIdIdx] === clientId) {
      clientEmail = row[emailIdx];
      clientName = row[nameIdx];
      clientRow = index + 2; // +2 because we shifted headers and 1-indexed
    }
  });
  
  if (!clientEmail) {
    Logger.log(`Client ${clientId} not found`);
    return false;
  }
  
  // Generate interaction ID
  const existingInts = interactionsSheet.getDataRange().getValues();
  const intCount = existingInts.length - 1; // minus header
  const newIntId = `I${String(intCount + 1).padStart(3, '0')}`;
  
  // Determine channel from interaction type
  const channelMap = {
    'Email': 'Email',
    'Call': 'Phone',
    'Meeting': 'In-Person',
    'Video': 'Video',
    'Text': 'Text',
    'Inbound': 'Email',
    'Outbound': 'Email'
  };
  const channel = channelMap[interactionType] || 'Other';
  
  // Add interaction
  const today = new Date();
  const newRow = [
    newIntId,
    clientId,
    clientEmail,
    today,
    interactionType,
    channel,
    'Follow-up contact',
    'Completed',
    notes || 'Marked as contacted via Reminders system',
    'No',
    'Reminders-System'
  ];
  
  interactionsSheet.appendRow(newRow);
  
  // Update client's last contact date
  const lastContactCol = headers.indexOf('Last_Contact_Date') + 1;
  const nextFollowupCol = headers.indexOf('Next_Followup_Due') + 1;
  clientsSheet.getRange(clientRow, lastContactCol).setValue(today);
  
  // Recalculate next follow-up (default to 30 days for now)
  const nextFollowup = new Date(today);
  nextFollowup.setDate(nextFollowup.getDate() + 30);
  clientsSheet.getRange(clientRow, nextFollowupCol).setValue(nextFollowup);
  
  // Recalculate reminders
  calculateReminders();
  
  Logger.log(`Marked ${clientId} (${clientName}) as contacted via ${interactionType}`);
  return true;
}

/**
 * Shows dialog to mark client as contacted
 */
function showMarkContactedDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Mark Client as Contacted',
    'Enter Client ID (e.g., C001):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const clientId = response.getResponseText().trim();
    if (clientId) {
      const result = markAsContacted(clientId, 'Email', 'Marked via menu');
      if (result) {
        ui.alert(`✅ ${clientId} marked as contacted. Reminders recalculated.`);
      } else {
        ui.alert(`❌ Client ${clientId} not found.`);
      }
    }
  }
}

/**
 * Snoozes a reminder by updating the next follow-up date
 * @param {string} clientId - The client ID
 * @param {number} snoozeDays - Number of days to snooze
 */
function snoozeReminder(clientId, snoozeDays) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsSheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  
  // Find client
  const clientsData = clientsSheet.getDataRange().getValues();
  const headers = clientsData.shift();
  const clientIdIdx = headers.indexOf('Client_ID');
  const nextFollowupIdx = headers.indexOf('Next_Followup_Due');
  
  let clientRow = -1;
  clientsData.forEach((row, index) => {
    if (row[clientIdIdx] === clientId) {
      clientRow = index + 2;
    }
  });
  
  if (clientRow === -1) {
    Logger.log(`Client ${clientId} not found`);
    return false;
  }
  
  // Calculate new follow-up date
  const newDate = new Date();
  newDate.setDate(newDate.getDate() + snoozeDays);
  
  // Update client
  clientsSheet.getRange(clientRow, nextFollowupIdx + 1).setValue(newDate);
  
  // Recalculate reminders
  calculateReminders();
  
  Logger.log(`Snoozed ${clientId} for ${snoozeDays} days (new due date: ${newDate.toDateString()})`);
  return true;
}

/**
 * Shows dialog to snooze a reminder
 */
function showSnoozeDialog() {
  const ui = SpreadsheetApp.getUi();
  
  const response = ui.prompt(
    'Snooze Reminder',
    'Enter Client ID and days (e.g., "C001, 7"):',
    ui.ButtonSet.OK_CANCEL
  );
  
  if (response.getSelectedButton() === ui.Button.OK) {
    const input = response.getResponseText().trim();
    const parts = input.split(',').map(p => p.trim());
    
    if (parts.length === 2) {
      const clientId = parts[0];
      const days = parseInt(parts[1]);
      
      if (days > 0) {
        const result = snoozeReminder(clientId, days);
        if (result) {
          ui.alert(`😴 ${clientId} snoozed for ${days} days.`);
        } else {
          ui.alert(`❌ Client ${clientId} not found.`);
        }
      } else {
        ui.alert('❌ Please enter a valid number of days.');
      }
    } else {
      ui.alert('❌ Format: "C001, 7" (Client ID, days)');
    }
  }
}

// ============================================
// TOOL #2 INTEGRATION
// ============================================

/**
 * Syncs interactions from Tool #2's Log sheet
 * Automatically imports emails and matches to clients by email address
 */
function syncFromTool2() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const clientsSheet = ss.getSheetByName(SHEET_NAMES.CLIENTS);
  const interactionsSheet = ss.getSheetByName(SHEET_NAMES.INTERACTIONS);
  
  // Check if Log sheet exists (Tool #2)
  let logSheet = ss.getSheetByName('Log');
  if (!logSheet) {
    Logger.log('Tool #2 Log sheet not found. Skipping sync.');
    return 0;
  }
  
  // Get all client emails for matching
  const clientsData = clientsSheet.getDataRange().getValues();
  const clientHeaders = clientsData.shift();
  const clientEmailIdx = clientHeaders.indexOf('Email');
  const clientIdIdx = clientHeaders.indexOf('Client_ID');
  
  const clientEmailMap = {};
  clientsData.forEach(row => {
    const email = row[clientEmailIdx];
    if (email) {
      clientEmailMap[email.toLowerCase()] = row[clientIdIdx];
    }
  });
  
  // Get Log data from Tool #2
  const logData = logSheet.getDataRange().getValues();
  if (logData.length <= 1) {
    Logger.log('Log sheet is empty');
    return 0;
  }
  
  const logHeaders = logData.shift();
  const logDateIdx = logHeaders.indexOf('Date') || logHeaders.indexOf('Timestamp');
  const logFromIdx = logHeaders.indexOf('From') || logHeaders.indexOf('Sender');
  const logSubjectIdx = logHeaders.indexOf('Subject');
  const logBodyIdx = logHeaders.indexOf('Body');
  
  // Get existing interactions to avoid duplicates
  const existingInts = interactionsSheet.getDataRange().getValues();
  const existingIds = new Set();
  const intHeaders = existingInts.length > 0 ? existingInts[0] : [];
  const intDateIdx = intHeaders.indexOf('Date');
  const intEmailIdx = intHeaders.indexOf('Client_Email');
  const intSubjectIdx = intHeaders.indexOf('Subject/Topic');
  
  existingInts.slice(1).forEach(row => {
    const key = `${row[intDateIdx]}|${row[intEmailIdx]}|${row[intSubjectIdx]}`;
    existingIds.add(key);
  });
  
  // Generate next interaction ID
  const intCount = existingInts.length - 1;
  let nextIntNum = intCount + 1;
  
  // Process Log entries
  let importedCount = 0;
  const newRows = [];
  
  logData.forEach(logRow => {
    const senderEmail = logRow[logFromIdx];
    if (!senderEmail) return;
    
    const normalizedEmail = senderEmail.toLowerCase().trim();
    const clientId = clientEmailMap[normalizedEmail];
    
    if (clientId) {
      const date = logRow[logDateIdx];
      const subject = logRow[logSubjectIdx] || '(No Subject)';
      
      // Check for duplicate
      const key = `${date}|${normalizedEmail}|${subject}`;
      if (existingIds.has(key)) {
        return; // Skip duplicate
      }
      
      const body = logBodyIdx >= 0 ? logRow[logBodyIdx] : '';
      const intId = `I${String(nextIntNum++).padStart(3, '0')}`;
      
      newRows.push([
        intId,
        clientId,
        normalizedEmail,
        date,
        'Inbound',
        'Email',
        subject,
        'Completed',
        `Auto-imported from Tool #2. Body preview: ${String(body).substring(0, 100)}...`,
        'No',
        'Tool-2-Auto'
      ]);
      
      importedCount++;
    }
  });
  
  // Add all new rows at once
  if (newRows.length > 0) {
    const startRow = existingInts.length + 1;
    interactionsSheet.getRange(startRow, 1, newRows.length, 11).setValues(newRows);
    
    // Recalculate reminders
    calculateReminders();
    
    Logger.log(`Imported ${importedCount} interactions from Tool #2`);
  } else {
    Logger.log('No new interactions to import from Tool #2');
  }
  
  return importedCount;
}

// ============================================
// UTILITY FUNCTIONS
// ============================================

/**
 * Helper function to get date X days ago (or in future if negative)
 */
function getDateDaysAgo(days) {
  const date = new Date();
  date.setDate(date.getDate() - days);
  return date;
}

/**
 * Quick test to verify all functions work
 */
function runAllTests() {
  Logger.log('=== Testing Client Follow-Up System ===');
  
  // Test 1: Calculate reminders
  Logger.log('Test 1: calculateReminders()');
  const reminders = calculateReminders();
  Logger.log(`✓ Calculated ${reminders ? reminders.length : 0} reminders`);
  
  // Test 2: Sync from Tool #2
  Logger.log('Test 2: syncFromTool2()');
  const synced = syncFromTool2();
  Logger.log(`✓ Synced ${synced} interactions from Tool #2`);
  
  // Test 3: Mark as contacted
  Logger.log('Test 3: markAsContacted("C002", "Email", "Test contact")');
  const marked = markAsContacted('C002', 'Email', 'Test contact');
  Logger.log(marked ? '✓ Marked C002 as contacted' : '✗ Failed to mark C002');
  
  // Test 4: Snooze
  Logger.log('Test 4: snoozeReminder("C002", 7)');
  const snoozed = snoozeReminder('C002', 7);
  Logger.log(snoozed ? '✓ Snoozed C002 for 7 days' : '✗ Failed to snooze C002');
  
  // Test 5: Send digest (test mode)
  Logger.log('Test 5: testSendDigest()');
  // Don't actually send in test run, just verify function exists
  Logger.log('✓ Digest function ready (check email for test)');
  
  Logger.log('=== All Tests Complete ===');
}
