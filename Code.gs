/**
 * @OnlyCurrentDoc
 * This script adds a custom menu to the spreadsheet for advanced scheduling.
 */

// Index (0-based) of the Settings sheet — third in the workbook.
const SETTINGS_SHEET_INDEX = 2;

// Global variables to hold settings.
let MAIN_PROTECTED_ROLES = [];
let IGNORED_ROLES_FOR_ASSIGNMENT = [];
let ROLE_EQUIVALENCIES = new Map();
let STATIC_ASSIGNMENTS = new Map();
let UNIQUE_ROLE_GROUPS = new Set();

/**
 * Creates a custom menu in the spreadsheet.
 */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Schedule Helper')
    .addItem('Fill Next Empty Meeting', 'fillNextEmptyMeetingColumn')
    .addToUi();
}

/**
 * Fetches all settings from the 'Settings' sheet.
 */
function fetchSettings() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const settingsSheet = ss.getSheets()[SETTINGS_SHEET_INDEX];
  if (!settingsSheet) {
    SpreadsheetApp.getUi().alert('Error: Could not find the Settings sheet (third sheet in the workbook).');
    return false;
  }

  const settingsData = settingsSheet.getDataRange().getDisplayValues();
  const headers = settingsData[0];
  
  // Clear previous settings
  MAIN_PROTECTED_ROLES = [];
  IGNORED_ROLES_FOR_ASSIGNMENT = [];
  ROLE_EQUIVALENCIES.clear();
  STATIC_ASSIGNMENTS.clear();
  UNIQUE_ROLE_GROUPS.clear();

  // Find column indices by header name for flexibility
  const protectedCol = headers.indexOf('Main Protected Roles');
  const ignoredCol = headers.indexOf('Ignored Roles for Assignment');
  const staticRoleCol = headers.indexOf('Static Role');
  const staticMemberCol = headers.indexOf('Assigned Member');

  for (let r = 1; r < settingsData.length; r++) {
    if (protectedCol > -1 && settingsData[r][protectedCol]) MAIN_PROTECTED_ROLES.push(settingsData[r][protectedCol]);
    if (ignoredCol > -1 && settingsData[r][ignoredCol]) IGNORED_ROLES_FOR_ASSIGNMENT.push(settingsData[r][ignoredCol]);
    if (staticRoleCol > -1 && staticMemberCol > -1 && settingsData[r][staticRoleCol] && settingsData[r][staticMemberCol]) {
        STATIC_ASSIGNMENTS.set(settingsData[r][staticRoleCol], settingsData[r][staticMemberCol]);
    }
  }

  // Fetch Role Equivalencies. All groups are now considered unique by default.
  for (let c = 0; c < headers.length; c++) {
    if (headers[c].toLowerCase().startsWith('equivalent roles')) {
        const group = [];
        for (let r = 1; r < settingsData.length; r++) {
            if (settingsData[r][c]) group.push(settingsData[r][c]);
        }
        if (group.length > 0) {
            const groupString = JSON.stringify(group.sort());
            group.forEach(role => ROLE_EQUIVALENCIES.set(role, groupString));
            UNIQUE_ROLE_GROUPS.add(groupString);
        }
    }
  }
  return true;
}

/**
 * Determines if a role requires a unique assignee for the current meeting.
 */
function isRoleProtected(roleName) {
    if (MAIN_PROTECTED_ROLES.includes(roleName)) return true;
    const group = ROLE_EQUIVALENCIES.get(roleName);
    if (group && UNIQUE_ROLE_GROUPS.has(group)) return true;
    return false;
}

/**
 * Finds the next meeting column with empty roles and fills them based on all rules.
 */
function fillNextEmptyMeetingColumn() {
  if (!fetchSettings()) return;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheets = ss.getSheets();
  const scheduleSheet = sheets[0];
  const availabilitySheet = sheets[1];

  const scheduleData = scheduleSheet.getDataRange().getDisplayValues();
  const allRoles = scheduleData.map(row => row[0]);
  const scheduleDates = scheduleData[0];
  
  let targetColIndex = -1;
  for (let c = 1; c < scheduleDates.length; c++) {
    for (let r = 1; r < allRoles.length; r++) {
      const roleName = allRoles[r];
      if (scheduleData[r][c] === '' && !IGNORED_ROLES_FOR_ASSIGNMENT.includes(roleName)) {
        targetColIndex = c;
        break;
      }
    }
    if (targetColIndex > -1) break;
  }

  if (targetColIndex === -1) {
    SpreadsheetApp.getUi().alert('No empty roles found to schedule!');
    return;
  }
  
  // --- PHASE 0: STATIC ASSIGNMENTS ---
  STATIC_ASSIGNMENTS.forEach((memberName, roleName) => {
      const rowIndex = allRoles.indexOf(roleName);
      if (rowIndex > -1) {
          scheduleSheet.getRange(rowIndex + 1, targetColIndex + 1).setValue(memberName);
      }
  });

  const alreadyAssignedThisTurn = new Set(STATIC_ASSIGNMENTS.values());

  const targetDate = scheduleDates[targetColIndex];
  const availabilityColIndex = availabilitySheet.getDataRange().getDisplayValues()[0].indexOf(targetDate);

  if (availabilityColIndex === -1) {
    SpreadsheetApp.getUi().alert(`Could not find the date "${targetDate}" in the Availability sheet.`);
    return;
  }

  const availabilityData = availabilitySheet.getDataRange().getDisplayValues();
  const availableMembers = [];
  for (let i = 1; i < availabilityData.length; i++) {
    if (availabilityData[i][availabilityColIndex] !== '0') {
      availableMembers.push(availabilityData[i][0]);
    }
  }
  
  const prevAssignments = new Map();
  if (targetColIndex > 1) {
    for (let r = 1; r < allRoles.length; r++) {
      const memberName = scheduleData[r][targetColIndex - 1];
      if (memberName) prevAssignments.set(memberName, allRoles[r]);
    }
  }

  const hadEquivalentRole = (member, currentRole) => {
    const lastRole = prevAssignments.get(member);
    if (!lastRole) return false;
    const currentGroup = ROLE_EQUIVALENCIES.get(currentRole);
    const lastGroup = ROLE_EQUIVALENCIES.get(lastRole);
    if (currentGroup && currentGroup === lastGroup) return true;
    return lastRole === currentRole;
  };
  
  const protectedToFill = [], otherToFill = [];
  const rolesData = scheduleSheet.getDataRange().getDisplayValues(); // Re-fetch data to see static assignments
  for (let r = 1; r < allRoles.length; r++) {
    const roleName = allRoles[r];
    if (rolesData[r][targetColIndex] === '' && !IGNORED_ROLES_FOR_ASSIGNMENT.includes(roleName)) {
        const cell = scheduleSheet.getRange(r + 1, targetColIndex + 1);
        if (isRoleProtected(roleName)) {
            protectedToFill.push({ cell, roleName });
        } else {
            otherToFill.push({ cell, roleName });
        }
    }
  }
  
  let candidatePool = availableMembers.filter(m => !alreadyAssignedThisTurn.has(m));
  for (let i = candidatePool.length - 1; i > 0; i--) {
      const j = Math.floor(Math.random() * (i + 1));
      [candidatePool[i], candidatePool[j]] = [candidatePool[j], candidatePool[i]];
  }

  // --- PHASE 1: FILL PROTECTED & UNIQUE GROUP ROLES ---
  protectedToFill.forEach(({ cell, roleName }) => {
      let candidateIndex = candidatePool.findIndex(member => !hadEquivalentRole(member, roleName));
      if (candidateIndex === -1 && candidatePool.length > 0) candidateIndex = 0; // Fallback

      if (candidateIndex !== -1) {
          const member = candidatePool[candidateIndex];
          cell.setValue(member);
          alreadyAssignedThisTurn.add(member);
          candidatePool.splice(candidateIndex, 1);
      } else {
          cell.setValue("NEEDS VOLUNTEER");
      }
  });

  // --- PHASE 2: FILL OTHER ROLES (WITH BETTER DISTRIBUTION) ---
  otherToFill.forEach(({cell, roleName}) => {
      let bestCandidate = null;

      // Tier 1: Prefer members who have NOT been assigned a role yet.
      let preferredPool = availableMembers.filter(m => !alreadyAssignedThisTurn.has(m));
      let idealPreferredPool = preferredPool.filter(m => !hadEquivalentRole(m, roleName));
      
      if (idealPreferredPool.length > 0) {
          bestCandidate = idealPreferredPool[Math.floor(Math.random() * idealPreferredPool.length)];
      } else {
          // Tier 2 (Fallback): Use any available member.
          let fallbackPool = availableMembers;
          let idealFallbackPool = fallbackPool.filter(m => !hadEquivalentRole(m, roleName));
          
          if (idealFallbackPool.length > 0) {
              bestCandidate = idealFallbackPool[Math.floor(Math.random() * idealFallbackPool.length)];
          } else if (availableMembers.length > 0) {
              // Super Fallback: If no one satisfies the "last week" rule, ignore it and pick anyone available.
              bestCandidate = availableMembers[Math.floor(Math.random() * availableMembers.length)];
          }
      }

      if (bestCandidate) {
          cell.setValue(bestCandidate);
          // IMPORTANT: Add the assigned person to the set so they are not prioritized for the NEXT non-protected role.
          alreadyAssignedThisTurn.add(bestCandidate);
      } else {
          cell.setValue("NEEDS VOLUNTEER");
      }
  });

  SpreadsheetApp.getUi().alert(`Successfully filled the schedule for ${targetDate}!`);
}