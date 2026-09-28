// ==UserScript==
// @name         Tech24 Case Timestamp Sync (Ultra-Fast Edition)
// @namespace    tech24-case-sla-sync
// @version      4.1
// @description  Instant loading, dynamic logging, flexible date formatting, and open-case handling.
// @match        https://tech24et.com/cases*
// @updateURL    https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-case-timestamp-sync.user.js
// @downloadURL  https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-case-timestamp-sync.user.js
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      *
// ==/UserScript==

(function () {
  "use strict";

  // ---- Config ----
  const BRIDGE_URL = "https://script.google.com/macros/s/AKfycbywn6bKxqdObHkgB4x7AgqyLLSCBYH2Mk5rSUfL_MndKiPfgeF8BRPd8o-TCGGOIbhI-g/exec";

  const SEARCH_SELECTOR = 'input[placeholder="Search Cases..."]';
  const FILTER_WAIT_TIMEOUT_MS = 12000;
  const FILTER_POLL_MS = 50;

  const WEEKDAY_NAMES = ["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"];
  let syncing = false;
  let cachedCols = null;
  let lastCheckTime = 0;

  // ---- Advanced UI & Error Manager ----
  const SyncUI = {
    container: null,
    msgEl: null,
    statusIcon: null,

    init() {
      if (document.getElementById('tech24-sync-widget')) return;

      this.container = document.createElement('div');
      this.container.id = 'tech24-sync-widget';
      this.container.style.cssText = `
        position: fixed; bottom: 24px; right: 24px; background: #1e293b; color: #f8fafc;
        padding: 12px 20px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        font-family: system-ui, sans-serif; z-index: 99999; display: flex;
        align-items: center; gap: 12px; transition: all 0.3s ease; border: 1px solid #334155;
        cursor: pointer; max-width: 400px;
      `;

      this.container.innerHTML = `
        <div id="tech24-sync-spinner" style="width: 12px; height: 12px; border-radius: 50%; background: #4ade80;"></div>
        <div style="flex-grow: 1;">
          <span id="tech24-sync-text" style="font-size: 14px; font-weight: 500; line-height: 1.4;">Sync: Idle (Asleep)</span>
        </div>
      `;

      this.container.onclick = () => {
        if (!syncing) checkAndRunSync(true);
      };

      document.body.appendChild(this.container);
      this.msgEl = document.getElementById('tech24-sync-text');
      this.statusIcon = document.getElementById('tech24-sync-spinner');
    },

    update(message, state = 'idle') {
      this.init();
      this.msgEl.innerHTML = message;

      if (state === 'idle') {
        this.statusIcon.style.background = '#94a3b8';
        this.statusIcon.style.animation = 'none';
        this.container.style.borderColor = '#334155';
      } else if (state === 'running') {
        this.statusIcon.style.background = '#3b82f6';
        this.statusIcon.style.animation = 'pulse 1.5s infinite';
        this.container.style.borderColor = '#3b82f6';
      } else if (state === 'success') {
        this.statusIcon.style.background = '#4ade80';
        this.statusIcon.style.animation = 'none';
        this.container.style.borderColor = '#4ade80';
      } else if (state === 'error') {
        this.statusIcon.style.background = '#ef4444';
        this.statusIcon.style.animation = 'none';
        this.container.style.borderColor = '#ef4444';
      }
    },

    showError(title, detail) {
      this.update(`<strong>${title}</strong><br><span style="font-size:12px; color:#cbd5e1;">${detail}</span>`, 'error');
    }
  };

  const style = document.createElement('style');
  style.innerHTML = `@keyframes pulse { 0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.7); } 70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(59, 130, 246, 0); } 100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); } }`;
  document.head.appendChild(style);

  // ---- Bridge / Network Calls ----
  function gmGetJson(url) {
    return new Promise((resolve, reject) => {
      GM_xmlhttpRequest({
        method: "GET",
        url: url,
        timeout: 45000,
        onload: (res) => {
          if (res.responseText.trim().toLowerCase().startsWith('<!doctype html>')) {
             if (res.status === 404) return reject(new Error("Dashboard cannot find Google Apps Script. Check URL."));
             if (res.status === 429) return reject(new Error("Google rate limit hit."));
             return reject(new Error(`Google server error HTML (Status: ${res.status}).`));
          }

          if (res.status >= 200 && res.status < 300) {
            try {
              const data = JSON.parse(res.responseText);
              if (data.error) return reject(new Error(`Sheet Error: ${data.error}`));
              resolve(data);
            } catch (e) { reject(new Error("Invalid JSON format received from Google.")); }
          } else {
            reject(new Error(`Server refused connection (Error ${res.status}).`));
          }
        },
        onerror: () => reject(new Error("Network disconnect. Please check your internet connection.")),
        ontimeout: () => reject(new Error("Google Servers took too long to respond."))
      });
    });
  }

  function bridgeGet(action, extraParams = "") {
    const query = extraParams ? `${extraParams}&_=${Date.now()}` : `_=${Date.now()}`;
    return gmGetJson(`${BRIDGE_URL}?action=${action}&${query}`);
  }

  // ---- Dashboard Data Extraction ----
  function getColumnIndices() {
    if (cachedCols) return cachedCols;
    const headerCells = Array.from(document.querySelectorAll("table thead th"));
    let found = { caseId: -1, start: -1, end: -1 };

    headerCells.forEach((th, idx) => {
      const label = th.textContent.toLowerCase().trim();
      if (label.includes("case id") || label === "caseid") found.caseId = idx;
      else if (label.includes("start") || label.includes("reg")) found.start = idx;
      else if (label.includes("end") || label.includes("close")) found.end = idx;
    });

    if (found.caseId === -1) found.caseId = 0;
    if (found.start === -1) found.start = 8;
    if (found.end === -1) found.end = 9;
    return (cachedCols = found);
  }

  function getSearchInput() {
    return document.querySelector(SEARCH_SELECTOR);
  }

  function setSearchValue(input, value) {
    const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value").set;
    nativeSetter.call(input, value);
    input.dispatchEvent(new Event("input", { bubbles: true }));
  }

  function findMatchingRow(caseId) {
    const cols = getColumnIndices();
    return Array.from(document.querySelectorAll("table tbody tr")).find((row) => {
      const cells = row.querySelectorAll("td");
      return cells.length && cells[cols.caseId]?.textContent.trim() === caseId;
    });
  }

  function waitForFilteredRow(caseId) {
    return new Promise((resolve) => {
      const start = Date.now();
      const check = () => {
        const row = findMatchingRow(caseId);
        if (row) {
          resolve(row);
          return;
        }
        if (Date.now() - start > FILTER_WAIT_TIMEOUT_MS) return resolve(null);
        setTimeout(check, FILTER_POLL_MS);
      };
      check();
    });
  }

  // Multi-format date parser to output "Mon 09-21-26"
  function formatDateForSheet(rawDate) {
    if (!rawDate) return "";
    const clean = rawDate.trim();

    let yyyy, mm, dd;

    // 1. Check YYYY-MM-DD or YYYY/MM/DD
    let match = clean.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})$/);
    if (match) {
      [, yyyy, mm, dd] = match;
    } else {
      // 2. Check MM-DD-YYYY or MM/DD/YYYY
      match = clean.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})$/);
      if (match) {
        [, mm, dd, yyyy] = match;
      } else {
        // 3. Fallback: Standard Date Object parsing
        const parsed = new Date(clean);
        if (!isNaN(parsed.getTime())) {
          yyyy = String(parsed.getFullYear());
          mm = String(parsed.getMonth() + 1);
          dd = String(parsed.getDate());
        } else {
          return clean;
        }
      }
    }

    const dateObj = new Date(parseInt(yyyy, 10), parseInt(mm, 10) - 1, parseInt(dd, 10));
    const dayName = WEEKDAY_NAMES[dateObj.getDay()];
    const formattedMM = String(mm).padStart(2, "0");
    const formattedDD = String(dd).padStart(2, "0");
    const formattedYY = String(yyyy).slice(-2);

    return `${dayName} ${formattedMM}-${formattedDD}-${formattedYY}`;
  }

  function splitTimestamp(text) {
    if (!text) return { date: null, time: null };
    const clean = text.replace(/^\(\vert{}\)$/g, "").trim();
    const lower = clean.toLowerCase();

    // Ignore empty/placeholder values for open/unclosed cases
    if (!clean || clean === "-" || lower.includes("pending") || lower.includes("open") || lower.includes("n/a")) {
      return { date: null, time: null };
    }

    const parts = clean.split(/\s+/);
    if (parts.length < 2) return { date: null, time: null };

    const dateStr = formatDateForSheet(parts[0]);
    const timeStr = parts.slice(1).join(" ");
    return { date: dateStr, time: timeStr };
  }

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  // ---- Main Sync Flow ----
  async function syncAll() {
    if (syncing) return;
    syncing = true;
    cachedCols = null;

    let stats = { inserted: 0, skipped: 0, leftBlank: 0 };

    try {
      SyncUI.update("Connecting to Google Sheets...", "running");

      const { rows } = await bridgeGet("getPendingCases");

      if (!rows || rows.length === 0) {
        await bridgeGet("clearTrigger", "summary={}");
        SyncUI.update("No empty rows found in sheet.", "idle");
        console.log("[Tech24 Sync] 📭 No pending cases found.");
        syncing = false;
        return;
      }

      console.log(`[Tech24 Sync] 🚀 Starting sync process... ${rows.length} rows ready to be filled.`);

      for (let i = 0; i < rows.length; i++) {
        const pending = rows[i];
        const input = getSearchInput();
        if (!input) throw new Error("Search bar not found on dashboard.");

        setSearchValue(input, pending.caseId);
        SyncUI.update(`Processing Case ${pending.caseId} (${i + 1}/${rows.length})...`, "running");

        const row = await waitForFilteredRow(pending.caseId);

        if (row) {
          const cols = getColumnIndices();
          const cells = row.querySelectorAll("td");
          const reg = splitTimestamp(cells[cols.start]?.textContent.trim());
          const closed = splitTimestamp(cells[cols.end]?.textContent.trim());

          const hasDataToInsert = (reg.date || closed.date);
          if (!hasDataToInsert) {
            stats.leftBlank++;
          } else {
            stats.inserted++;
          }

          const params = new URLSearchParams({
            row: String(pending.row),
            regDate: reg.date || "",
            regTime: reg.time || "",
            closedDate: closed.date || "",
            closedTime: closed.time || ""
          });

          for (let attempt = 1; attempt <= 3; attempt++) {
            try {
              await bridgeGet("submitResult", params.toString());
              console.log(`[Tech24 Sync] ✅ Updated case ${pending.caseId} (Start: ${reg.date || 'None'}, Closed: ${closed.date || 'Open'})`);
              break;
            } catch (submitErr) {
              console.warn(`[Tech24 Sync] ⚠️ Retry on case ${pending.caseId} (Attempt ${attempt}/3): ${submitErr.message}`);
              if (attempt < 3) await delay(1500);
            }
          }

        } else {
          stats.skipped++;
          console.log(`[Tech24 Sync] ⏭️ Case not found on dashboard after 12s, skipped: ${pending.caseId}`);
        }
      }

      await bridgeGet("clearTrigger", `summary=${JSON.stringify(stats)}`);
      SyncUI.update(`Sync Complete! Inserted: ${stats.inserted}`, "success");

      console.log(`[Tech24 Sync] ✅ Finished inserting timestamps!`);
      console.log(`[Tech24 Sync] 📊 Summary - Inserted: ${stats.inserted} | Skipped: ${stats.skipped} | Left Blank: ${stats.leftBlank}`);

      setTimeout(() => SyncUI.update("Sync: Idle (Asleep)", "idle"), 5000);

    } catch (err) {
      console.error("[Tech24 Sync Error] Critical Sync Failure:", err);
      SyncUI.showError("Sync Failed", err.message);
    } finally {
      syncing = false;
    }
  }

  // ---- Zero-Idle Event Listeners ----
  async function checkAndRunSync(manualTrigger = false) {
    if (syncing) return;

    if (!manualTrigger && (Date.now() - lastCheckTime < 2000)) return;
    lastCheckTime = Date.now();

    try {
      if (manualTrigger) SyncUI.update("Checking for trigger...", "running");
      const { shouldRun } = await bridgeGet("checkTrigger");

      if (shouldRun) {
        await syncAll();
      } else if (manualTrigger) {
        SyncUI.update("No active trigger in Sheets.", "idle");
        setTimeout(() => SyncUI.update("Sync: Idle (Asleep)", "idle"), 3000);
      }
    } catch (err) {
      SyncUI.showError("Connection Issue", err.message);
    }
  }

  window.addEventListener('focus', () => checkAndRunSync(false));
  document.addEventListener('visibilitychange', () => {
    if (document.visibilityState === 'visible') checkAndRunSync(false);
  });

  setTimeout(() => {
    SyncUI.init();
    checkAndRunSync(false);
  }, 1000);

})();
