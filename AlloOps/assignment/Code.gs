const ALLOOPS = {
  sheets: {
    config: 'Config',
    tasks: 'Tasks',
    cycles: 'Cycles',
    managerDashboard: 'Manager Dashboard',
    events: 'Events',
    notifications: 'Notifications'
  },
  headers: {
    config: ['Key', 'Value', 'Description'],
    tasks: [
      'Task ID', 'Enabled', 'Task Name', 'Assignee Email', 'Frequency (days)',
      'Due Time (HH:mm)', 'Proof Type', 'Proof Instructions', 'Manager Email',
      'Leadership Email(s)', 'Manager Ack SLA Hours', 'Escalation SLA Hours',
      'Start Date'
    ],
    cycles: [
      'Cycle ID', 'Task ID', 'Task Name', 'Assignee Email', 'Scheduled Date',
      'Due Date', 'Due At', 'Proof Type', 'Proof Instructions', 'Status',
      'Completion Evidence', 'Completed By', 'Completed At',
      'Manager Acknowledged', 'Acknowledged By', 'Acknowledged At',
      'Manager Alert Sent At', 'Escalated At', 'Last Updated At',
      'Manager Email', 'Leadership Email(s)', 'Manager Ack SLA Hours',
      'Escalation SLA Hours'
    ],
    events: ['Timestamp', 'Actor Email', 'Event Type', 'Cycle ID', 'Details'],
    notifications: [
      'Notification ID', 'Cycle ID', 'Type', 'Recipient(s)', 'Status',
      'Created At', 'Sent At', 'Error'
    ]
  },
  defaults: [
    ['TIME_ZONE', '', 'IANA time zone; initialized from the spreadsheet.'],
    ['WORKING_DAYS', '1,2,3,4,5', 'ISO weekdays: Monday=1 through Sunday=7.'],
    ['WORKDAY_START', '09:00', 'Start of the working day, 24-hour time.'],
    ['WORKDAY_END', '17:00', 'End of the working day, 24-hour time.'],
    ['HOLIDAYS', '', 'Comma-separated dates in yyyy-MM-dd format.'],
    ['GENERATION_HORIZON_DAYS', '30', 'How far ahead to create upcoming task cycles.']
  ]
};

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('AlloOps')
    .addItem('Set up task tracker', 'setupTracker')
    .addItem('Open My Tasks', 'showTaskSidebar')
    .addItem('Generate task cycles', 'generateTaskCycles')
    .addItem('Run alerts now', 'processAutomation')
    .addToUi();
}

function setupTracker() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  ensureSheet_(ss, ALLOOPS.sheets.config, ALLOOPS.headers.config);
  ensureSheet_(ss, ALLOOPS.sheets.tasks, ALLOOPS.headers.tasks);
  ensureSheet_(ss, ALLOOPS.sheets.cycles, ALLOOPS.headers.cycles);
  ensureSheet_(ss, ALLOOPS.sheets.managerDashboard, ['AlloOps Manager Dashboard']);
  ensureSheet_(ss, ALLOOPS.sheets.events, ALLOOPS.headers.events);
  ensureSheet_(ss, ALLOOPS.sheets.notifications, ALLOOPS.headers.notifications);
  initializeConfig_(ss);
  seedSampleTasks_(ss);
  initializeManagerDashboard_(ss);
  ensureAutomationTrigger_();
  showTaskSidebar();
}

function showTaskSidebar() {
  SpreadsheetApp.getUi().showSidebar(
    HtmlService.createHtmlOutputFromFile('Sidebar').setTitle('AlloOps task tracker')
  );
}

function generateTaskCycles() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another task-cycle operation is running. Try again shortly.');
  try {
    const created = generateTaskCycles_();
    SpreadsheetApp.getUi().alert('Task cycle generation completed. New cycles: ' + created);
  } finally {
    lock.releaseLock();
  }
}

function processAutomation() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another automation run is active. Try again shortly.');
  try {
    generateTaskCycles_();
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const config = readConfig_(ss);
    const cyclesSheet = getSheet_(ss, ALLOOPS.sheets.cycles);
    const cycles = readRows_(cyclesSheet, ALLOOPS.headers.cycles);
    const notifications = {};
    readRows_(getSheet_(ss, ALLOOPS.sheets.notifications), ALLOOPS.headers.notifications).forEach(function (row) {
      notifications[String(row['Notification ID'])] = row;
    });
    const now = new Date();

    cycles.forEach(function (cycle) {
      if (String(cycle.Status).toUpperCase() === 'CANCELLED') return;
      if (!cycle['Due At']) throw new Error('Missing Due At for cycle ' + cycle['Cycle ID'] + '.');
      const dueAt = parseSheetDate_(cycle['Due At'], 'Due At for cycle ' + cycle['Cycle ID']);
      const complete = String(cycle.Status).toUpperCase() === 'COMPLETED';
      let alertSentAt = cycle['Manager Alert Sent At']
        ? parseSheetDate_(cycle['Manager Alert Sent At'], 'Manager Alert Sent At for cycle ' + cycle['Cycle ID'])
        : null;

      if (!complete && !alertSentAt && now > dueAt && isWorkingDay_(now, config)) {
        const recipients = normalizeEmail_(cycle['Manager Email']);
        const notificationId = String(cycle['Cycle ID']) + ':MANAGER_OVERDUE';
        const previous = notifications[notificationId];
        if (previous && String(previous.Status).toUpperCase() === 'SENT' && previous['Sent At']) {
          alertSentAt = parseSheetDate_(previous['Sent At'], 'Sent At for notification ' + notificationId);
          updateCycle_(cyclesSheet, cycle, {
            'Manager Alert Sent At': alertSentAt,
            'Last Updated At': now
          });
        } else if (!previous) {
          const sent = sendNotificationOnce_(
            ss, cycle, 'MANAGER_OVERDUE', recipients,
            'Overdue task: ' + cycle['Task Name'],
            buildOverdueMessage_(cycle, config),
            notifications
          );
          if (sent) {
            alertSentAt = now;
            updateCycle_(cyclesSheet, cycle, {
              'Manager Alert Sent At': now,
              'Last Updated At': now
            });
            appendEvent_(ss, '', 'MANAGER_ALERT_SENT', cycle['Cycle ID'], 'Overdue notification sent to ' + recipients);
          }
        }
      }

      if (!alertSentAt || cycle['Escalated At']) return;
      const ackSla = positiveNumber_(cycle['Manager Ack SLA Hours'], 'Manager Ack SLA Hours', cycle['Task ID']);
      const escalationSla = positiveNumber_(cycle['Escalation SLA Hours'], 'Escalation SLA Hours', cycle['Task ID']);
      const ackDue = addBusinessHours_(alertSentAt, ackSla, config);
      const escalationDue = addBusinessHours_(alertSentAt, escalationSla, config);
      const unacknowledged = !isTruthy_(cycle['Manager Acknowledged']) && now >= ackDue;
      const unresolved = !complete && now >= escalationDue;
      if (!unacknowledged && !unresolved) return;

      const leadership = normalizeRecipients_(cycle['Leadership Email(s)']);
      if (!leadership) throw new Error('Leadership Email(s) is required to escalate cycle ' + cycle['Cycle ID'] + '.');
      const reasons = [];
      if (unresolved) reasons.push('the task is still incomplete');
      if (unacknowledged) reasons.push('the manager alert has not been acknowledged');
      const notificationId = String(cycle['Cycle ID']) + ':LEADERSHIP_ESCALATION';
      const previous = notifications[notificationId];
      if (previous && String(previous.Status).toUpperCase() === 'SENT') {
        updateCycle_(cyclesSheet, cycle, {
          'Escalated At': previous['Sent At'] ? parseSheetDate_(previous['Sent At'], 'Sent At for notification ' + notificationId) : now,
          'Last Updated At': now
        });
      } else if (!previous) {
        const sent = sendNotificationOnce_(
          ss, cycle, 'LEADERSHIP_ESCALATION', leadership,
          'Task escalation: ' + cycle['Task Name'],
          buildEscalationMessage_(cycle, reasons),
          notifications
        );
        if (sent) {
          updateCycle_(cyclesSheet, cycle, {
            'Escalated At': now,
            'Last Updated At': now
          });
          appendEvent_(ss, '', 'LEADERSHIP_ESCALATED', cycle['Cycle ID'], reasons.join('; ') + '; sent to ' + leadership);
        }
      }
    });
  } finally {
    lock.releaseLock();
  }
}

function getSidebarData() {
  const email = requireActiveEmail_();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(ss);
  const today = formatDate_(new Date(), 'yyyy-MM-dd', config);
  const pending = [];
  const alerts = [];
  const teamTasks = [];
  readRows_(getSheet_(ss, ALLOOPS.sheets.cycles), ALLOOPS.headers.cycles).forEach(function (cycle) {
    if (String(cycle.Status).toUpperCase() === 'CANCELLED') return;
    const dueAt = parseSheetDate_(cycle['Due At'], 'Due At for cycle ' + cycle['Cycle ID']);
    const scheduledDate = formatDate_(
      parseSheetDate_(cycle['Scheduled Date'], 'Scheduled Date for cycle ' + cycle['Cycle ID']),
      'yyyy-MM-dd',
      config
    );
    const item = {
      id: String(cycle['Cycle ID']),
      name: String(cycle['Task Name']),
      due: formatDate_(dueAt, 'yyyy-MM-dd HH:mm'),
      instructions: String(cycle['Proof Instructions'] || ''),
      proofType: String(cycle['Proof Type'])
    };
    if (String(cycle.Status).toUpperCase() !== 'COMPLETED' &&
        scheduledDate <= today &&
        normalizeEmail_(cycle['Assignee Email']) === email) {
      pending.push(item);
    }
    if (String(cycle.Status).toUpperCase() === 'PENDING' &&
        scheduledDate <= today &&
        normalizeEmail_(cycle['Manager Email']) === email) {
      teamTasks.push({
        id: item.id,
        name: item.name,
        assignee: String(cycle['Assignee Email']),
        due: item.due,
        status: new Date() > dueAt ? 'OVERDUE' : 'PENDING',
        proofType: item.proofType
      });
    }
    if (cycle['Manager Alert Sent At'] &&
        !isTruthy_(cycle['Manager Acknowledged']) &&
        normalizeEmail_(cycle['Manager Email']) === email) {
      alerts.push(item);
    }
  });
  return { email: email, pending: pending, teamTasks: teamTasks, alerts: alerts };
}

function submitCompletion(cycleId, evidence) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another task update is running. Try again shortly.');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = requireActiveEmail_();
    const cyclesSheet = getSheet_(ss, ALLOOPS.sheets.cycles);
    const cycle = findCycle_(cyclesSheet, cycleId);
    if (!cycle) throw new Error('Task cycle not found. Refresh the task list and try again.');
    if (normalizeEmail_(cycle['Assignee Email']) !== email) {
      throw new Error('Only the assigned employee can submit completion evidence for this task.');
    }
    const config = readConfig_(ss);
    const today = formatDate_(new Date(), 'yyyy-MM-dd', config);
    const scheduledDate = formatDate_(
      parseSheetDate_(cycle['Scheduled Date'], 'Scheduled Date for cycle ' + cycleId),
      'yyyy-MM-dd',
      config
    );
    if (scheduledDate > today) {
      throw new Error('This task is not available until its scheduled date (' + scheduledDate + ').');
    }
    if (String(cycle.Status).toUpperCase() === 'COMPLETED') {
      throw new Error('This task cycle is already completed; duplicate submissions are not accepted.');
    }
    const proof = validateEvidence_(String(cycle['Proof Type']), String(evidence || '').trim());
    const now = new Date();
    updateCycle_(cyclesSheet, cycle, {
      Status: 'COMPLETED',
      'Completion Evidence': proof,
      'Completed By': email,
      'Completed At': now,
      'Last Updated At': now
    });
    appendEvent_(ss, email, 'TASK_COMPLETED', String(cycleId), 'Proof type: ' + cycle['Proof Type']);
    generateTaskCycles_();
    return 'Completion recorded for ' + cycle['Task Name'] + '.';
  } finally {
    lock.releaseLock();
  }
}

function acknowledgeAlert(cycleId) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) throw new Error('Another task update is running. Try again shortly.');
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const email = requireActiveEmail_();
    const cyclesSheet = getSheet_(ss, ALLOOPS.sheets.cycles);
    const cycle = findCycle_(cyclesSheet, cycleId);
    if (!cycle) throw new Error('Task cycle not found. Refresh the alert list and try again.');
    if (normalizeEmail_(cycle['Manager Email']) !== email) {
      throw new Error('Only the configured manager can acknowledge this alert.');
    }
    if (!cycle['Manager Alert Sent At']) throw new Error('This cycle has no manager alert to acknowledge.');
    if (isTruthy_(cycle['Manager Acknowledged'])) return 'This manager alert was already acknowledged.';
    const now = new Date();
    updateCycle_(cyclesSheet, cycle, {
      'Manager Acknowledged': true,
      'Acknowledged By': email,
      'Acknowledged At': now,
      'Last Updated At': now
    });
    appendEvent_(ss, email, 'MANAGER_ALERT_ACKNOWLEDGED', String(cycleId), 'Task status was not changed.');
    return 'Manager alert acknowledged.';
  } finally {
    lock.releaseLock();
  }
}

function generateTaskCycles_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const config = readConfig_(ss);
  const taskRows = readRows_(getSheet_(ss, ALLOOPS.sheets.tasks), ALLOOPS.headers.tasks);
  const cyclesSheet = getSheet_(ss, ALLOOPS.sheets.cycles);
  const allCycles = readRows_(cyclesSheet, ALLOOPS.headers.cycles);
  const cyclesByTask = {};
  allCycles.forEach(function (cycle) {
    const taskId = String(cycle['Task ID']);
    if (!cyclesByTask[taskId]) cyclesByTask[taskId] = [];
    cyclesByTask[taskId].push(cycle);
  });
  const horizon = Number(config.GENERATION_HORIZON_DAYS);
  if (!Number.isInteger(horizon) || horizon < 1 || horizon > 365) {
    throw new Error('GENERATION_HORIZON_DAYS must be an integer from 1 through 365.');
  }
  const today = formatDate_(new Date(), 'yyyy-MM-dd');
  const through = addDateDays_(today, horizon);
  const tasks = [];
  const taskIds = new Set();
  taskRows.forEach(function (row) {
    if (!isEnabled_(row.Enabled)) return;
    const task = validateTask_(row, config);
    if (taskIds.has(task.id)) throw new Error('Duplicate Task ID: ' + task.id + '.');
    taskIds.add(task.id);
    tasks.push(task);
  });

  const newCycles = [];
  const newEvents = [];
  let reactivated = 0;
  tasks.forEach(function (task) {
    const taskCycles = cyclesByTask[task.id] || [];
    const completed = taskCycles.filter(function (cycle) {
      return String(cycle.Status).toUpperCase() === 'COMPLETED';
    });
    const pending = taskCycles.filter(function (cycle) {
      return String(cycle.Status).toUpperCase() === 'PENDING';
    });
    let scheduled;
    let dueDate;

    if (completed.length) {
      const latestCompleted = completed.reduce(function (latest, cycle) {
        return !latest || parseSheetDate_(cycle['Completed At'], 'Completed At').getTime() >
          parseSheetDate_(latest['Completed At'], 'Completed At').getTime() ? cycle : latest;
      }, null);
      const completedDate = formatDate_(
        parseSheetDate_(latestCompleted['Completed At'], 'Completed At'),
        'yyyy-MM-dd',
        config
      );
      scheduled = addDateDays_(completedDate, task.frequency);
      if (task.frequency === 1) scheduled = nextWorkingDate_(scheduled, config);
      dueDate = nextWorkingDate_(scheduled, config);
      const expectedId = task.id + '_' + scheduled.replace(/-/g, '');
      const keep = pending.find(function (cycle) { return String(cycle['Cycle ID']) === expectedId; }) || null;
      cancelLegacyCycles_(ss, cyclesSheet, pending.filter(function (cycle) { return cycle !== keep; }));
      if (keep) return;
    } else if (pending.length) {
      const firstPending = pending.reduce(function (first, cycle) {
        return !first || parseSheetDate_(cycle['Scheduled Date'], 'Scheduled Date').getTime() <
          parseSheetDate_(first['Scheduled Date'], 'Scheduled Date').getTime() ? cycle : first;
      }, null);
      cancelLegacyCycles_(ss, cyclesSheet, pending.filter(function (cycle) { return cycle !== firstPending; }));
      return;
    } else {
      scheduled = task.startDate;
      if (task.frequency === 1) scheduled = nextWorkingDate_(scheduled, config);
      if (scheduled > through) return;
      dueDate = nextWorkingDate_(scheduled, config);
    }

    const id = task.id + '_' + scheduled.replace(/-/g, '');
    const now = new Date();
    const existingCycle = taskCycles.find(function (cycle) { return String(cycle['Cycle ID']) === id; });
    if (existingCycle) {
      if (String(existingCycle.Status).toUpperCase() === 'CANCELLED') {
        updateCycle_(cyclesSheet, existingCycle, {
          'Task ID': task.id,
          'Task Name': task.name,
          'Assignee Email': task.assignee,
          'Scheduled Date': dateAtTime_(scheduled, '00:00', config),
          'Due Date': dateAtTime_(dueDate, '00:00', config),
          'Due At': dateAtTime_(dueDate, task.dueTime, config),
          'Proof Type': task.proofType,
          'Proof Instructions': task.instructions,
          Status: 'PENDING',
          'Completion Evidence': '',
          'Completed By': '',
          'Completed At': '',
          'Manager Acknowledged': false,
          'Acknowledged By': '',
          'Acknowledged At': '',
          'Manager Alert Sent At': '',
          'Escalated At': '',
          'Last Updated At': now,
          'Manager Email': task.manager,
          'Leadership Email(s)': task.leadership,
          'Manager Ack SLA Hours': task.ackSla,
          'Escalation SLA Hours': task.escalationSla
        });
        appendEvent_(ss, '', 'CYCLE_REACTIVATED', id, 'Reused an existing logical cycle ID for completion-based recurrence.');
        reactivated++;
      }
      return;
    }
    const newCycle = [
      id, task.id, task.name, task.assignee, dateAtTime_(scheduled, '00:00', config),
      dateAtTime_(dueDate, '00:00', config), dateAtTime_(dueDate, task.dueTime, config),
      task.proofType, task.instructions, 'PENDING', '', '', '', false, '', '', '',
      '', now, task.manager, task.leadership, task.ackSla, task.escalationSla
    ];
    newCycles.push(newCycle);
    newEvents.push([now, '', 'CYCLE_CREATED', id, 'Scheduled for ' + dueDate]);
    taskCycles.push({
      'Cycle ID': id,
      'Task ID': task.id,
      Status: 'PENDING',
      'Scheduled Date': dateAtTime_(scheduled, '00:00', config)
    });
    cyclesByTask[task.id] = taskCycles;
  });
  if (newCycles.length) {
    cyclesSheet.getRange(cyclesSheet.getLastRow() + 1, 1, newCycles.length, ALLOOPS.headers.cycles.length)
      .setValues(newCycles);
    const eventsSheet = getSheet_(ss, ALLOOPS.sheets.events);
    eventsSheet.getRange(eventsSheet.getLastRow() + 1, 1, newEvents.length, ALLOOPS.headers.events.length)
      .setValues(newEvents);
  }
  return newCycles.length + reactivated;
}

function cancelLegacyCycles_(ss, cyclesSheet, cycles) {
  cycles.forEach(function (cycle) {
    updateCycle_(cyclesSheet, cycle, {
      Status: 'CANCELLED',
      'Last Updated At': new Date()
    });
    appendEvent_(
      ss,
      '',
      'LEGACY_CYCLE_CANCELLED',
      String(cycle['Cycle ID']),
      'Superseded by completion-anchored recurrence; original record retained.'
    );
  });
}

function initializeConfig_(ss) {
  const sheet = getSheet_(ss, ALLOOPS.sheets.config);
  if (sheet.getLastRow() > 1) return;
  const defaults = ALLOOPS.defaults.map(function (row) {
    return row[0] === 'TIME_ZONE'
      ? [row[0], ss.getSpreadsheetTimeZone(), row[2]]
      : row;
  });
  sheet.getRange(2, 1, defaults.length, 3).setValues(defaults);
}

function seedSampleTasks_(ss) {
  const sheet = getSheet_(ss, ALLOOPS.sheets.tasks);
  if (sheet.getLastRow() > 1) return;
  const config = readConfig_(ss);
  const startDate = dateAtTime_(formatDate_(new Date(), 'yyyy-MM-dd', config), '00:00', config);
  const manager = 'manager@example.com';
  const leadership = 'leadership@example.com';
  const samples = [
    [
      'DEMO_DAILY', false, 'Daily Operations Update', 'employee.a@example.com', 1, '17:00',
      'DAILY_CONFIRMATION', 'Confirm the daily checks completed and note any exceptions.',
      manager, leadership, 4, 16, startDate
    ],
    [
      'DEMO_7D', false, 'Portal Task A', 'employee.b@example.com', 7, '17:00',
      'TICKET_ID', 'Complete the portal task and enter its Ticket ID.',
      manager, leadership, 4, 16, startDate
    ],
    [
      'DEMO_15D', false, 'Portal Task B', 'employee.c@example.com', 15, '17:00',
      'TICKET_ID', 'Complete the portal task and enter its Ticket ID.',
      manager, leadership, 4, 16, startDate
    ],
    [
      'DEMO_30D', false, 'Portal Task C', 'employee.d@example.com', 30, '17:00',
      'TICKET_ID', 'Complete the portal task and enter its Ticket ID.',
      manager, leadership, 4, 16, startDate
    ]
  ];
  sheet.getRange(2, 1, samples.length, ALLOOPS.headers.tasks.length).setValues(samples);
}

function initializeManagerDashboard_(ss) {
  const sheet = getSheet_(ss, ALLOOPS.sheets.managerDashboard);
  sheet.getRange('A1').setValue('AlloOps Manager Dashboard');
  sheet.getRange('A2').setValue(
    'Live view of pending and overdue cycles. Acknowledge manager alerts in AlloOps > Open My Tasks.'
  );
  sheet.getRange('A3:A6').setValues([
    ['Open task cycles'],
    ['Overdue task cycles'],
    ['Alerts awaiting acknowledgement'],
    ['Escalated cycles']
  ]);
  sheet.getRange('B3:B6').setFormulas([
    ['=COUNTIF(Cycles!J2:J,"PENDING")'],
    ['=COUNTIFS(Cycles!J2:J,"PENDING",Cycles!G2:G,"<"&NOW())'],
    ['=COUNTIFS(Cycles!J2:J,"PENDING",Cycles!Q2:Q,"<>",Cycles!N2:N,FALSE)'],
    ['=COUNTIFS(Cycles!R2:R,"<>",Cycles!J2:J,"<>CANCELLED")']
  ]);
  sheet.getRange('A8:J8').setValues([[
    'Cycle ID', 'Task', 'Assignee', 'Due At', 'Required Proof', 'Proof Instructions',
    'Status', 'Manager Acknowledged', 'Manager Alert Sent At', 'Escalated At'
  ]]);
  sheet.getRange('A9').setFormula(
    '=IFERROR(FILTER({Cycles!A2:A,Cycles!C2:C,Cycles!D2:D,Cycles!G2:G,Cycles!H2:H,' +
    'Cycles!I2:I,IF(Cycles!G2:G<NOW(),"OVERDUE","PENDING"),Cycles!N2:N,Cycles!Q2:Q,Cycles!R2:R},' +
    'Cycles!J2:J="PENDING"),"No open tasks")'
  );
  sheet.setFrozenRows(8);
  sheet.getRange('A1:J1').setFontWeight('bold').setFontSize(14);
  sheet.getRange('A8:J8').setFontWeight('bold');
  sheet.getRange('D9:D').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.getRange('I9:J').setNumberFormat('yyyy-mm-dd hh:mm');
  sheet.autoResizeColumns(1, 10);
}

function ensureSheet_(ss, name, headers) {
  let sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    return;
  }
  const actual = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0].map(String);
  if (headers.some(function (header, index) { return actual[index] !== header; })) {
    throw new Error('The ' + name + ' sheet has unexpected columns. Expected: ' + headers.join(' | '));
  }
}

function ensureAutomationTrigger_() {
  const exists = ScriptApp.getProjectTriggers().some(function (trigger) {
    return trigger.getHandlerFunction() === 'processAutomation';
  });
  if (!exists) ScriptApp.newTrigger('processAutomation').timeBased().everyHours(1).create();
}

function readConfig_(ss) {
  const config = {};
  readRows_(getSheet_(ss, ALLOOPS.sheets.config), ALLOOPS.headers.config).forEach(function (row) {
    config[String(row.Key)] = String(row.Value == null ? '' : row.Value).trim();
  });
  if (!config.TIME_ZONE) config.TIME_ZONE = ss.getSpreadsheetTimeZone();
  return config;
}

function readRows_(sheet, headers) {
  if (sheet.getLastRow() < 2) return [];
  return sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).getValues()
    .map(function (values) {
      const row = {};
      headers.forEach(function (header, index) { row[header] = values[index]; });
      return row;
    })
    .filter(function (row) {
      return headers.some(function (header) { return row[header] !== '' && row[header] != null; });
    });
}

function validateTask_(row, config) {
  const id = String(row['Task ID'] || '').trim();
  if (!id || !String(row['Task Name'] || '').trim()) throw new Error('Enabled tasks need a Task ID and Task Name.');
  const assignee = normalizeEmail_(row['Assignee Email']);
  const manager = normalizeEmail_(row['Manager Email']);
  const leadership = normalizeRecipients_(row['Leadership Email(s)']);
  if (!isEmail_(assignee)) throw new Error('Invalid Assignee Email for task ' + id + '.');
  if (!isEmail_(manager)) throw new Error('Invalid Manager Email for task ' + id + '.');
  if (!leadership || leadership.split(',').some(function (email) { return !isEmail_(email); })) {
    throw new Error('Valid Leadership Email(s) are required for task ' + id + '.');
  }
  const frequency = Number(row['Frequency (days)']);
  if ([1, 7, 15, 30].indexOf(frequency) === -1) {
    throw new Error('Frequency for task ' + id + ' must be 1, 7, 15, or 30.');
  }
  const dueTime = normalizeTime_(row['Due Time (HH:mm)'], 'Due Time for task ' + id, config);
  const proofType = String(row['Proof Type'] || '').trim().toUpperCase();
  if (['TICKET_ID', 'DAILY_CONFIRMATION', 'TEXT'].indexOf(proofType) === -1) {
    throw new Error('Proof Type for task ' + id + ' must be TICKET_ID, DAILY_CONFIRMATION, or TEXT.');
  }
  const instructions = String(row['Proof Instructions'] || '').trim();
  if (!instructions) throw new Error('Proof Instructions are required for task ' + id + '.');
  const ackSla = positiveNumber_(row['Manager Ack SLA Hours'], 'Manager Ack SLA Hours', id);
  const escalationSla = positiveNumber_(row['Escalation SLA Hours'], 'Escalation SLA Hours', id);
  const startDate = normalizeDate_(row['Start Date'], 'Start Date for task ' + id, config);
  return {
    id: id,
    name: String(row['Task Name']).trim(),
    assignee: assignee,
    manager: manager,
    leadership: leadership,
    frequency: frequency,
    dueTime: dueTime,
    proofType: proofType,
    instructions: instructions,
    ackSla: ackSla,
    escalationSla: escalationSla,
    startDate: startDate
  };
}

function validateEvidence_(proofType, evidence) {
  if (!evidence || /[\u0000-\u0008\u000B\u000C\u000E-\u001F]/.test(evidence)) {
    throw new Error('Completion evidence is required and must not contain control characters.');
  }
  if (proofType === 'TICKET_ID' &&
      (evidence.length > 100 || !/[A-Za-z0-9]/.test(evidence) || /[\r\n\t]/.test(evidence))) {
    throw new Error('Enter a valid portal Ticket ID (up to 100 characters).');
  }
  if (proofType === 'DAILY_CONFIRMATION' && evidence.length < 10) {
    throw new Error('Daily confirmation must describe checks completed (at least 10 characters).');
  }
  if (proofType === 'TEXT' && evidence.length > 1000) {
    throw new Error('Text evidence must be 1,000 characters or fewer.');
  }
  if (['TICKET_ID', 'DAILY_CONFIRMATION', 'TEXT'].indexOf(proofType) === -1) {
    throw new Error('Unsupported evidence type: ' + proofType + '.');
  }
  return evidence;
}

function sendNotificationOnce_(ss, cycle, type, recipients, subject, body, notifications) {
  const sheet = getSheet_(ss, ALLOOPS.sheets.notifications);
  const notificationId = String(cycle['Cycle ID']) + ':' + type;
  if (notifications[notificationId]) return false;
  const row = sheet.getLastRow() + 1;
  const createdAt = new Date();
  sheet.getRange(row, 1, 1, ALLOOPS.headers.notifications.length)
    .setValues([[notificationId, cycle['Cycle ID'], type, recipients, 'RESERVED', createdAt, '', '']]);
  notifications[notificationId] = {
    'Notification ID': notificationId,
    Status: 'RESERVED',
    'Created At': createdAt
  };
  SpreadsheetApp.flush();
  try {
    MailApp.sendEmail({ to: recipients, subject: subject, body: body });
    const sentAt = new Date();
    sheet.getRange(row, 5, 1, 4).setValues([['SENT', createdAt, sentAt, '']]);
    notifications[notificationId].Status = 'SENT';
    notifications[notificationId]['Sent At'] = sentAt;
    return true;
  } catch (error) {
    sheet.getRange(row, 5, 1, 4).setValues([['FAILED', createdAt, '', String(error.message || error)]]);
    throw error;
  }
}

function appendEvent_(ss, actor, eventType, cycleId, details) {
  getSheet_(ss, ALLOOPS.sheets.events).appendRow([new Date(), actor, eventType, cycleId, details]);
}

function updateCycle_(sheet, cycle, updates) {
  const rowNumber = findRowNumber_(sheet, 'Cycle ID', String(cycle['Cycle ID']));
  if (!rowNumber) throw new Error('Task cycle disappeared during update: ' + cycle['Cycle ID']);
  Object.keys(updates).forEach(function (key) {
    const column = ALLOOPS.headers.cycles.indexOf(key) + 1;
    if (!column) throw new Error('Unknown cycle field: ' + key);
    sheet.getRange(rowNumber, column).setValue(updates[key]);
  });
}

function findCycle_(sheet, cycleId) {
  return readRows_(sheet, ALLOOPS.headers.cycles).find(function (row) {
    return String(row['Cycle ID']) === String(cycleId);
  }) || null;
}

function findRowNumber_(sheet, header, value) {
  const column = ALLOOPS.headers.cycles.indexOf(header) + 1;
  if (sheet.getLastRow() < 2) return 0;
  const values = sheet.getRange(2, column, sheet.getLastRow() - 1, 1).getValues();
  for (let index = 0; index < values.length; index++) {
    if (String(values[index][0]) === value) return index + 2;
  }
  return 0;
}

function getSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing required sheet "' + name + '". Choose AlloOps > Set up task tracker.');
  return sheet;
}

function requireActiveEmail_() {
  const email = normalizeEmail_(Session.getActiveUser().getEmail());
  if (!isEmail_(email)) {
    throw new Error('Google did not provide your signed-in email. Use a Workspace account with identity sharing enabled.');
  }
  return email;
}

function normalizeEmail_(value) {
  return String(value || '').trim().toLowerCase();
}

function normalizeRecipients_(value) {
  return String(value || '').split(/[;,]/).map(function (email) {
    return email.trim().toLowerCase();
  }).filter(Boolean).join(',');
}

function isEmail_(value) {
  return /^[^@\s]+@[^@\s]+\.[^@\s]+$/.test(value);
}

function isEnabled_(value) {
  return value === true || String(value).trim().toUpperCase() === 'TRUE' || String(value).trim() === '1';
}

function isTruthy_(value) {
  return value === true || String(value).trim().toUpperCase() === 'TRUE';
}

function positiveNumber_(value, label, taskId) {
  const number = Number(value);
  if (!Number.isFinite(number) || number <= 0) {
    throw new Error(label + ' for task ' + taskId + ' must be a positive number of hours.');
  }
  return number;
}

function parseSheetDate_(value, label) {
  const date = value instanceof Date ? new Date(value.getTime()) : new Date(value);
  if (isNaN(date.getTime())) throw new Error(label + ' must be a valid date/time.');
  return date;
}

function normalizeDate_(value, label, config) {
  if (!value) throw new Error(label + ' is required.');
  if (value instanceof Date) return formatDate_(value, 'yyyy-MM-dd', config);
  const text = String(value).trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(text)) throw new Error(label + ' must be a valid date.');
  const parts = text.split('-').map(Number);
  const parsed = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2]));
  if (parsed.toISOString().slice(0, 10) !== text) throw new Error(label + ' must be a valid date.');
  return text;
}

function normalizeTime_(value, label, config) {
  const text = value instanceof Date
    ? formatDate_(value, 'HH:mm', config)
    : String(value || '').trim();
  if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(text)) {
    throw new Error(label + ' must use HH:mm (24-hour) format.');
  }
  return text;
}

function formatDate_(date, pattern, config) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const timezone = config && config.TIME_ZONE ? config.TIME_ZONE : ss.getSpreadsheetTimeZone();
  return Utilities.formatDate(date, timezone, pattern);
}

function dateAtTime_(dateText, timeText, config) {
  const dateParts = dateText.split('-').map(Number);
  const timeParts = timeText.split(':').map(Number);
  const localAsUtc = Date.UTC(dateParts[0], dateParts[1] - 1, dateParts[2], timeParts[0], timeParts[1]);
  let candidate = localAsUtc;
  for (let attempt = 0; attempt < 3; attempt++) {
    const formatted = Utilities.formatDate(new Date(candidate), config.TIME_ZONE, "yyyy-MM-dd'T'HH:mm");
    const parts = formatted.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2})$/);
    const representedAsUtc = Date.UTC(
      Number(parts[1]), Number(parts[2]) - 1, Number(parts[3]), Number(parts[4]), Number(parts[5])
    );
    const offset = representedAsUtc - candidate;
    const adjusted = localAsUtc - offset;
    if (adjusted === candidate) break;
    candidate = adjusted;
  }
  return new Date(candidate);
}

function addDateDays_(dateText, count) {
  const parts = dateText.split('-').map(Number);
  const date = new Date(Date.UTC(parts[0], parts[1] - 1, parts[2] + count));
  return date.toISOString().slice(0, 10);
}

function isWorkingDay_(date, config) {
  const weekday = Number(Utilities.formatDate(date, config.TIME_ZONE, 'u'));
  const workingDays = String(config.WORKING_DAYS || '1,2,3,4,5').split(',').map(Number);
  const dayText = formatDate_(date, 'yyyy-MM-dd', config);
  const holidays = String(config.HOLIDAYS || '').split(',').map(function (day) { return day.trim(); });
  if (!workingDays.length || workingDays.some(function (day) { return day < 1 || day > 7 || !Number.isInteger(day); })) {
    throw new Error('WORKING_DAYS must contain ISO weekday numbers from 1 through 7.');
  }
  holidays.forEach(function (holiday) {
    if (holiday && !/^\d{4}-\d{2}-\d{2}$/.test(holiday)) {
      throw new Error('HOLIDAYS entries must use yyyy-MM-dd format.');
    }
  });
  return workingDays.indexOf(weekday) !== -1 && holidays.indexOf(dayText) === -1;
}

function nextWorkingDate_(dateText, config) {
  let candidate = dateText;
  for (let attempts = 0; attempts <= 14; attempts++) {
    if (isWorkingDay_(dateAtTime_(candidate, '00:00', config), config)) return candidate;
    candidate = addDateDays_(candidate, 1);
  }
  throw new Error('No working day found within 14 days. Check WORKING_DAYS and HOLIDAYS.');
}

function addBusinessHours_(start, hours, config) {
  let cursor = start;
  let remaining = hours * 60;
  const opening = normalizeTime_(config.WORKDAY_START || '09:00', 'WORKDAY_START', config);
  const closing = normalizeTime_(config.WORKDAY_END || '17:00', 'WORKDAY_END', config);
  const openParts = opening.split(':').map(Number);
  const closeParts = closing.split(':').map(Number);
  const openMinute = openParts[0] * 60 + openParts[1];
  const closeMinute = closeParts[0] * 60 + closeParts[1];
  if (closeMinute <= openMinute) throw new Error('WORKDAY_END must be later than WORKDAY_START.');

  while (remaining > 0) {
    const dateText = formatDate_(cursor, 'yyyy-MM-dd', config);
    if (!isWorkingDay_(cursor, config)) {
      cursor = dateAtTime_(nextWorkingDate_(dateText, config), opening, config);
      continue;
    }
    const currentTime = formatDate_(cursor, 'HH:mm', config).split(':').map(Number);
    let currentMinute = currentTime[0] * 60 + currentTime[1];
    if (currentMinute < openMinute) {
      cursor = dateAtTime_(dateText, opening, config);
      currentMinute = openMinute;
    }
    if (currentMinute >= closeMinute) {
      cursor = dateAtTime_(nextWorkingDate_(addDateDays_(dateText, 1), config), opening, config);
      continue;
    }
    const consumed = Math.min(closeMinute - currentMinute, remaining);
    cursor = new Date(cursor.getTime() + consumed * 60000);
    remaining -= consumed;
    if (remaining > 0) {
      const nextDate = nextWorkingDate_(addDateDays_(formatDate_(cursor, 'yyyy-MM-dd', config), 1), config);
      cursor = dateAtTime_(nextDate, opening, config);
    }
  }
  return cursor;
}

function buildOverdueMessage_(cycle, config) {
  return [
    'Task: ' + cycle['Task Name'],
    'Assignee: ' + cycle['Assignee Email'],
    'Due: ' + formatDate_(parseSheetDate_(cycle['Due At'], 'Due At'), 'yyyy-MM-dd HH:mm', config),
    'Cycle: ' + cycle['Cycle ID'],
    'Proof required: ' + cycle['Proof Type'] + ' — ' + cycle['Proof Instructions'],
    '',
    'Open the AlloOps Google Sheet, review the task cycle, and acknowledge this alert in the sidebar.'
  ].join('\n');
}

function buildEscalationMessage_(cycle, reasons) {
  return [
    'Task: ' + cycle['Task Name'],
    'Assignee: ' + cycle['Assignee Email'],
    'Manager: ' + cycle['Manager Email'],
    'Cycle: ' + cycle['Cycle ID'],
    'Status: ' + cycle.Status,
    'Escalation reason(s): ' + reasons.join('; '),
    'Task completion and manager acknowledgement are tracked separately in the AlloOps sheet.'
  ].join('\n');
}
