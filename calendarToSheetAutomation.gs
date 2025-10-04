const CONFIG = {
  SPREADSHEET_ID: '1e-71aCV0YUTU_Wr4--RGu_SUePYlmmCIfBAyMX97xsI',
  SHEET_NAME:      'Trainer Sheet',
  CALENDAR_ID:     'primary',
  START_HOUR:      19,   // only 7 PM+
  END_HOUR:        21,   // before 9 PM
  DAYS_BEHIND:     30,    // look 7 days back
  DAYS_AHEAD:      30    // look 30 days forward
};

/**
 * Main entrypoint: fetch events, parse trainer/bootcamp, write into sheet.
 */
function runCalendarSheetsAutomation() {
  try {
    console.log('▶️ Starting Calendar→Sheet automation');
    
    // 1) fetch calendar & sheet
    const cal   = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
    const ss    = SpreadsheetApp.openById(CONFIG.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Sheet not found: ${CONFIG.SHEET_NAME}`);
    
    // 2) compute date window
    const now  = new Date();
    const from = new Date(now);
    from.setDate(now.getDate() - CONFIG.DAYS_BEHIND);
    const to   = new Date(now);
    to.setDate(now.getDate() + CONFIG.DAYS_AHEAD);
    
    // 3) fetch & filter events by hour
    const allEvents = cal.getEvents(from, to);
    const events = allEvents.filter(ev => {
      const h = ev.getStartTime().getHours();
      // If event is on Saturday (6), use 12–14; otherwise use CONFIG range
      if (ev.getStartTime().getDay() === 6) { // Saturday
        return h >= 12 && h <= 14;
      } else {
        return h >= CONFIG.START_HOUR && h < CONFIG.END_HOUR;
      }
    });
    console.log(`→ ${events.length}/${allEvents.length} events between ${CONFIG.START_HOUR}:00–${CONFIG.END_HOUR}:00`);
    
    // 4) pull sheet data + header
    const data    = sheet.getDataRange().getValues();
    const headers = sheet
      .getRange(1, 1, 1, sheet.getLastColumn())
      .getValues()[0];
    
    const codePatterns = [
  // 1) FALLBACK: any “Data Science Bootcamp” without a number → DS10
  {
    re: /\bData\s+Science\s+Bootcamp\b(?!.*\d)/i,
    name: 'DS10'
  },
      
      // 1) Explicit numbered forms (so they win first)
      { re: /\bData\s+Science\s+Bootcamp\s*11\b/i,         name: 'DS11'       },
      { re: /\bData\s+Science\s+Bootcamp\s*12\b/i,         name: 'DS12'       },
      { re: /\bData\s+Science\s+Bootcamp\s*13\b/i,           name: 'DS13'       },

      { re: /\bData\s+Analytics\s+Bootcamp\s*05\b/i,       name: 'DA 05'      },
      { re: /\bData\s+Analytics\s+Bootcamp\s*06\b.*White\b/i, name: 'DA6 White'  },
      { re: /\bData\s+Analytics\s+Bootcamp\s*06\b.*Black\b/i, name: 'DA6 Black'  },
      { re: /\bData\s+Analytics\s+Bootcamp\s*7\b.*Blue\b/i, name: 'DA7 Blue'  },
      { re: /\bData\s+Analytics\s+Bootcamp\s*7\b.*Green\b/i, name: 'DA7 Green'  },

      { re: /\bAgentic\s*AI\b/i,                          name: 'Agentic AI 02' },
      { re: /\bAgentic\s*AI\b/i,                          name: 'Agentic AI 02' },

      // 2) FALLBACK generic (only if none of the above matched)
      { re: /\bData\s+Analytics\s+Bootcamp\s*06\b/i,      name: 'DA6 White'  },

      // 3) Very‐short codes (last resort)
      { re: /\bDS\s*11\b/i,                               name: 'DS11'       },
      { re: /\bDS\s*12\b/i,                               name: 'DS12'       },
      { re: /\bDS\s*13\b/i,                               name: 'DS13'       },
      { re: /\bDA\s*05\b/i,                               name: 'DA 05'      },
      { re: /\bDA\s*6\s*White\b/i,                        name: 'DA6 White'  },
      { re: /\bDA\s*6\s*Black\b/i,                        name: 'DA6 Black'  },
      { re: /\bDA\s*4\b/i,                                name: 'DA4'        },
      { re: /\bDS\s*13\b/i,                               name: 'DS13'       },
      { re: /\bDA\s*7\s*Blue\b/i,                         name: 'DA7 Blue'  },
      { re: /\bDA\s*7\s*Green\b/i,                        name: 'DA7 Green'  },
      
    ];


    let updatesCount = 0;
    
    // 6) process each event independently
    events.forEach(event => {
      try {
        const title   = event.getTitle()       || '';
        const desc    = event.getDescription() || '';
        const evDate  = event.getStartTime();
        const text    = `${title} ${desc}`.replace(/\s+/g,' ').trim();
        
        // -- extract trainer name (robust) --
function extractTrainer(text, event) {
  const patterns = [
    // Titles: Mr/Ms/Mrs/Miss/Dr + multi-word name
    /(?:^|\b)(Mr|Ms|Mrs|Miss|Dr)\.?\s+([A-Z][\w'.-]+(?:\s+[A-Z][\w'.-]+){0,3})/i,

    // with/led by/hosted by/facilitated by/session with + name
    /(?:with|led by|hosted by|facilitated by|session with)\s+([A-Z][\w'.-]+(?:\s+[A-Z][\w'.-]+){0,3})/i,

    // trainer/instructor/teacher: Name  OR “trainer - Name”
    /(?:trainer|instructor|teacher)\s*[:\-–]\s*([A-Z][\w'.-]+(?:\s+[A-Z][\w'.-]+){0,3})/i,

    // (trainer: Name)  or  (instructor Name)
    /\((?:trainer|instructor|teacher)\s*[:\-–]?\s*([A-Z][\w'.-]+(?:\s+[A-Z][\w'.-]+){0,3})\)/i,

    // “… Name … trainer …” (name preceding the word trainer)
    /\b([A-Z][\w'.-]+(?:\s+[A-Z][\w'.-]+)+)\b(?=[^)]*\btrainer\b)/i
  ];

  for (const re of patterns) {
    const m = text.match(re);
    if (m) return (m[2] || m[1]).trim();
  }

  // Fallback: use first non-organizer guest’s display name/email local-part
  try {
    const guests = event.getGuestList().filter(g => !g.isOrganizer());
    if (guests.length) {
      const g = guests[0];
      const guess = (g.getName() || g.getEmail().split('@')[0])
        .replace(/[._]+/g, ' ')
        .trim();
      if (guess) return guess;
    }
  } catch (e) { /* guest list may be unavailable; ignore */ }

  return null;
}

const trainer = extractTrainer(text, event);
        
        // -- extract bootcamp code via regex list --
        let bc = null;
        for (const {re,name} of codePatterns) {
          if (re.test(text)) {
            bc = name;
            break;
          }
        }
        
        console.log(`➡️ "${title}" → trainer: ${trainer}, bootcamp: ${bc}, date: ${evDate.toDateString()}`);
        if (!trainer || !bc) throw new Error('Missing trainer or bootcamp');
        
        // -- find column index --
        const colIdx = headers.findIndex(h =>
          h && h.toString().trim().toLowerCase() === bc.toLowerCase()
        );
        if (colIdx < 0) {
          console.warn(`⚠️ No column for "${bc}". Available headers: ${headers.join(' | ')}`);
          return;
        }
        const col = colIdx + 1;  // Sheets API is 1-based
        
        // -- find matching row by Y/M/D only --
        let rowIdx = -1;
        for (let i = 1; i < data.length; i++) {
          const cell = data[i][1];  // column B
          if (cell instanceof Date) {
            if (
              cell.getFullYear() === evDate.getFullYear() &&
              cell.getMonth()    === evDate.getMonth() &&
              cell.getDate()     === evDate.getDate()
            ) { rowIdx = i; break; }
          } else if (typeof cell === 'string') {
            const pd = new Date(cell);
            if (!isNaN(pd) &&
                pd.getFullYear() === evDate.getFullYear() &&
                pd.getMonth()    === evDate.getMonth() &&
                pd.getDate()     === evDate.getDate()
            ) { rowIdx = i; break; }
          }
        }
        if (rowIdx < 0) {
          console.warn(`⚠️ No row for date ${evDate.toDateString()}`);
          return;
        }
        
        // annotate the trainer with duration before writing
          const annotated = annotateWithDuration(event, trainer);
          sheet.getRange(rowIdx + 1, col).setValue(annotated);

        
      } catch (evtErr) {
        console.error(`✗ Skipping event "${event.getTitle()}": ${evtErr.message}`);
      }
    });
    
    console.log(`✅ Finished: ${updatesCount} cells updated.`);
    if (updatesCount) sendNotificationEmail(updatesCount);
    
  } catch (err) {
    console.error('🚨 Fatal error:', err);
    sendErrorNotificationEmail(err.toString());
  }
}

/**
 * Annotates a trainer name with the event’s duration.
 * e.g. “Alice (1 hour)” or “Bob (1.5 hour)”
 */
function annotateWithDuration(event, trainer) {
  const start  = event.getStartTime();
  const end    = event.getEndTime();
  const diffMs = end.getTime() - start.getTime();
  let hours    = diffMs / (1000 * 60 * 60);

  // Round to the nearest half-hour
  hours = Math.round(hours * 2) / 2;

  // Build display string (drops “.0” automatically)
  const display = (hours % 1 === 0) ? `${hours}` : `${hours}`;
  return `${trainer} (${display} hour)`;
}


/**
 * Send notification when updates occur
 */
function sendNotificationEmail(updatesCount) {
  const subject = `Calendar→Sheet: ${updatesCount} updates`;
  const body    = `The automation wrote ${updatesCount} trainer entries.`;
  const email   = Session.getActiveUser().getEmail();
  if (email) MailApp.sendEmail(email, subject, body);
}

/**
 * Send notification on fatal error
 */
function sendErrorNotificationEmail(errorMessage) {
  const subject = 'Calendar→Sheet Automation Error';
  const body    = `An error occurred:\n\n${errorMessage}`;
  const email   = Session.getActiveUser().getEmail();
  if (email) MailApp.sendEmail(email, subject, body);
}


/**
 * (Optional) Test that time filtering works
 */
function testTimeFiltering() {
  const cal = CalendarApp.getCalendarById(CONFIG.CALENDAR_ID);
  const now = new Date();
  const from = new Date(now); from.setDate(now.getDate() - CONFIG.DAYS_BEHIND);
  const to   = new Date(now); to.setDate(now.getDate() + CONFIG.DAYS_AHEAD);
  const evs  = cal.getEvents(from, to)
                 .filter(e => {
                   const h = e.getStartTime().getHours();
                   return h >= CONFIG.START_HOUR && h < CONFIG.END_HOUR;
                 });
  console.log(`Found ${evs.length} events between ${CONFIG.START_HOUR}:00–${CONFIG.END_HOUR}:00`);
}


/**
 * (Optional) Create a daily trigger at 22:00
 */
function setupDailyTriggerWithTimeFilter() {
  // delete old
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runCalendarSheetsAutomation')
    .forEach(t => ScriptApp.deleteTrigger(t));
  // schedule new
  ScriptApp.newTrigger('runCalendarSheetsAutomation')
    .timeBased()
    .everyDays(1)
    .atHour(22)
    .create();
  console.log('Daily trigger set at 22:00');
}

/**
 * (Optional) Remove any existing automation trigger
 */
function removeTriggers() {
  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'runCalendarSheetsAutomation')
    .forEach(t => ScriptApp.deleteTrigger(t));
  console.log('All Calendar→Sheet triggers removed');
}