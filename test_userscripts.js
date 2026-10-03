const assert = require("assert");
const fs = require("fs");
const path = require("path");

// Helper to simulate browser window environment for testing
function createBrowserContext() {
  const window = {
    addEventListener() {}
  };
  const document = {
    head: { appendChild() {} },
    body: { appendChild() {} },
    createElement() {
      return {
        style: {},
        setAttribute() {},
        appendChild() {},
        dispatchEvent() {}
      };
    },
    getElementById() { return null; },
    querySelector() { return null; },
    querySelectorAll() { return []; },
    addEventListener() {}
  };
  global.window = window;
  global.document = document;
  global.KeyboardEvent = class {};
  global.Event = class {};
  global.MutationObserver = class { observe() {} };
  global.GM_getValue = () => null;
  global.GM_setValue = () => {};
}

createBrowserContext();

// Load userscripts from new tampermonkey scripts directory or fallback
const syncPath = path.resolve(__dirname, "./weekly report scripts/new scripts/tampermonkey scripts/tech24-case-timestamp-sync.user.js");
const slaPath = path.resolve(__dirname, "./weekly report scripts/new scripts/tampermonkey scripts/tech24-sla-duration.user.js");

if (fs.existsSync(syncPath)) {
  require(syncPath);
} else {
  require("./tech24-case-timestamp-sync.user.js");
}

if (fs.existsSync(slaPath)) {
  require(slaPath);
} else {
  require("./tech24-sla-duration.user.js");
}

const SyncScript = global.window.Tech24TimestampSync;
const SLAScript = global.window.Tech24SLADuration;

console.log("🧪 Running unit tests for Tech24 Userscripts...\n");

// --- Test 1: DateParser formatting in Timestamp Sync ---
{
  console.log("Testing Timestamp Sync - DateParser...");
  const dateFormatted = SyncScript.DateParser.formatDateForSheet("2026-09-21");
  assert.strictEqual(dateFormatted, "Mon 09-21-26", `Expected Mon 09-21-26, got: ${dateFormatted}`);

  const splitResult = SyncScript.DateParser.splitTimestamp("2026-09-21 02:30 PM");
  assert.strictEqual(splitResult.date, "Mon 09-21-26");
  assert.strictEqual(splitResult.time, "02:30 PM");

  const openCaseResult = SyncScript.DateParser.splitTimestamp("-");
  assert.strictEqual(openCaseResult.date, null);
  assert.strictEqual(openCaseResult.time, null);

  console.log("  ✅ Timestamp Sync DateParser passed!");
}

// --- Test 2: SLA & Duration - Date Parsing ---
{
  console.log("Testing SLA Duration - DateParser...");
  const parsedDate = SLAScript.DateParser.parseDashboardDate("(2026-08-14 01:51 PM)");
  assert.ok(parsedDate instanceof Date);
  assert.strictEqual(parsedDate.getFullYear(), 2026);
  assert.strictEqual(parsedDate.getMonth(), 7); // August = index 7
  assert.strictEqual(parsedDate.getDate(), 14);
  assert.strictEqual(parsedDate.getHours(), 13);
  assert.strictEqual(parsedDate.getMinutes(), 51);

  const formattedStr = SLAScript.DateParser.formatDateTime(parsedDate);
  assert.strictEqual(formattedStr, "14-08-2026 01:51 PM");

  const durationStr = SLAScript.DateParser.formatDuration(135);
  assert.strictEqual(durationStr, "2h 15m");

  console.log("  ✅ SLA Duration DateParser passed!");
}

// --- Test 3: SLA Business Hours Calculation ---
{
  console.log("Testing SLA Duration - BusinessCalendar...");
  // Mon 8:00 AM start, Zone 1 (7 business hours target -> Mon 3:00 PM)
  const startMon = new Date(2026, 8, 21, 8, 0, 0); // Mon Sep 21 2026
  const targetMon = SLAScript.BusinessCalendar.addBusinessHours(startMon, 7);
  assert.strictEqual(targetMon.getHours(), 15);
  assert.strictEqual(targetMon.getDate(), 21);

  // Fri 3:00 PM start, Zone 1 (7 business hours target -> 2h Fri + 5h Sat -> Sat 1:00 PM)
  const startFri = new Date(2026, 8, 25, 15, 0, 0); // Fri Sep 25 2026
  const targetSat = SLAScript.BusinessCalendar.addBusinessHours(startFri, 7);
  assert.strictEqual(targetSat.getDay(), 6); // Saturday
  assert.strictEqual(targetSat.getHours(), 13); // 1:00 PM

  // Business minutes diff
  const diffMinutes = SLAScript.BusinessCalendar.businessMinutesDiff(startFri, targetSat);
  assert.strictEqual(diffMinutes, 420); // 7 hours = 420 mins

  console.log("  ✅ SLA Duration BusinessCalendar passed!");
}

// --- Test 4: Header Column Matching ---
{
  console.log("Testing Header Matching...");
  const mockTable = {
    querySelectorAll(selector) {
      if (selector === "thead th") {
        return [
          { textContent: '<th class="py-2 px-1 whitespace-nowrap"><div class="flex gap-2 items-center">Case ID</div></th>' },
          { textContent: 'Status' },
          { textContent: '<th class="py-2 px-1 whitespace-nowrap"><div class="flex gap-2 items-center">Start</div></th>' },
          { textContent: '<th class="py-2 px-1 whitespace-nowrap"><div class="flex gap-2 items-center">End</div></th>' },
          { textContent: 'District' }
        ];
      }
      return [];
    }
  };

  const caseIdx = SLAScript.TableRenderer.getColumnIndex(mockTable, "case id");
  const startIdx = SLAScript.TableRenderer.getColumnIndex(mockTable, "start");
  const endIdx = SLAScript.TableRenderer.getColumnIndex(mockTable, "end");
  const districtIdx = SLAScript.TableRenderer.getColumnIndex(mockTable, "district");

  assert.strictEqual(caseIdx, 0);
  assert.strictEqual(startIdx, 2);
  assert.strictEqual(endIdx, 3);
  assert.strictEqual(districtIdx, 4);

  console.log("  ✅ Header Matching passed!");
}

// --- Test 5: Google Sheets Script Duration Calculation Engine ---
{
  console.log("Testing Google Sheets Duration Calculation...");
  const gasScriptContent = fs.readFileSync(path.resolve(__dirname, "google-apps-script-bridge.gs"), "utf8");

  // Mock GAS environment
  global.SpreadsheetApp = {
    openById() {
      return {
        getSheetByName() {
          return {
            getLastRow() { return 10; },
            getRange() {
              return {
                getValues() {
                  // Sheet Col B (idx 1) = Machine ID, Sheet Col H (idx 7) = Zone
                  return [
                    ["", "M001", "", "", "", "", "", "1"],
                    ["", "M002", "", "", "", "", "", "2"],
                    ["", "M003", "", "", "", "", "", "3"]
                  ];
                }
              };
            }
          };
        }
      };
    }
  };

  // Evaluate GAS calculation functions in sandbox context
  const contextFunc = new Function("global", `
    ${gasScriptContent}
    return {
      parseDateTime,
      formatTargetDateTime,
      formatDurationMs,
      calculateRowDurationData,
      getZoneForMachineId
    };
  `);

  const gasEngine = contextFunc(global);

  // Test 5a: Zone Lookup
  assert.strictEqual(gasEngine.getZoneForMachineId("M001"), 1);
  assert.strictEqual(gasEngine.getZoneForMachineId("M002"), 2);
  assert.strictEqual(gasEngine.getZoneForMachineId("M003"), 3);

  // Test 5b: Standard Duration Calculation & Target Format
  // Start: 05-10-2026 06:20 AM, End: 05-10-2026 08:05 AM (1h 45m duration)
  // Zone 1 = 7 hours target -> Tgt: 05-10-2026 01:20 PM
  const regDate = "05-10-2026";
  const regTime = "06:20 AM";
  const closedDate = "05-10-2026";
  const closedTime = "08:05 AM";

  const result1 = gasEngine.calculateRowDurationData("M001", regDate, regTime, closedDate, closedTime);
  assert.ok(result1);
  assert.strictEqual(result1.durationStr, "1h 45m");
  assert.strictEqual(result1.zone, 1);
  assert.strictEqual(result1.targetTimeStr, "05-10-2026 01:20 PM");
  assert.strictEqual(result1.isOverdue, false);
  assert.strictEqual(result1.outputString, "1h 45m | Z1 | Tgt: 05-10-2026 01:20 PM");

  // Test 5c: Overdue SLA Highlighting Check
  // Start: 05-10-2026 06:20 AM, Zone 1 (7h target -> 1:20 PM). End: 05-10-2026 02:00 PM (Overdue)
  const resultOverdue = gasEngine.calculateRowDurationData("M001", regDate, regTime, closedDate, "02:00 PM");
  assert.ok(resultOverdue);
  assert.strictEqual(resultOverdue.isOverdue, true);

  console.log("  ✅ Google Sheets Duration Calculation Engine passed!");
}

console.log("\n🎉 All unit tests passed successfully!");
