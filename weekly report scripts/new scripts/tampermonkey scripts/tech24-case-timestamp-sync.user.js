// ==UserScript==
// @name         Tech24 Case Timestamp Sync (Ultra-Fast Edition)
// @namespace    tech24-case-sla-sync
// @version      4.2
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

  // =========================================================================
  // 1. Configuration & Constants
  // =========================================================================
  const CONFIG = Object.freeze({
    BRIDGE_URL: "https://script.google.com/macros/s/AKfycbywn6bKxqdObHkgB4x7AgqyLLSCBYH2Mk5rSUfL_MndKiPfgeF8BRPd8o-TCGGOIbhI-g/exec",
    SEARCH_SELECTOR: 'input[placeholder="Search Cases..."]',
    FILTER_WAIT_TIMEOUT_MS: 12000,
    FILTER_POLL_MS: 50,
    CHECK_INTERVAL_MS: 2000,
    MAX_RETRIES: 3,
    RETRY_DELAY_MS: 1500,
    WEEKDAY_NAMES: Object.freeze(["Sun", "Mon", "Tue", "Wed", "Thu", "Fri", "Sat"])
  });

  let syncing = false;
  let cachedCols = null;
  let lastCheckTime = 0;

  // =========================================================================
  // 2. UI Manager Component
  // =========================================================================
  const SyncUI = {
    container: null,
    msgEl: null,
    statusIcon: null,

    init() {
      if (document.getElementById("tech24-sync-widget")) return;

      this.injectStyles();

      this.container = document.createElement("div");
      this.container.id = "tech24-sync-widget";
      this.container.style.cssText = `
        position: fixed; bottom: 24px; right: 24px; background: #1e293b; color: #f8fafc;
        padding: 12px 20px; border-radius: 8px; box-shadow: 0 4px 15px rgba(0,0,0,0.3);
        font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        z-index: 99999; display: flex; align-items: center; gap: 12px; transition: all 0.3s ease;
        border: 1px solid #334155; cursor: pointer; max-width: 400px; user-select: none;
      `;

      this.container.innerHTML = `
        <div id="tech24-sync-spinner" style="width: 12px; height: 12px; border-radius: 50%; background: #94a3b8; flex-shrink: 0;"></div>
        <div style="flex-grow: 1;">
          <span id="tech24-sync-text" style="font-size: 14px; font-weight: 500; line-height: 1.4;">Sync: Idle (Asleep)</span>
        </div>
      `;

      this.container.onclick = () => {
        if (!syncing) checkAndRunSync(true);
      };

      document.body.appendChild(this.container);
      this.msgEl = document.getElementById("tech24-sync-text");
      this.statusIcon = document.getElementById("tech24-sync-spinner");
    },

    injectStyles() {
      if (document.getElementById("tech24-sync-style")) return;
      const style = document.createElement("style");
      style.id = "tech24-sync-style";
      style.textContent = `
        @keyframes tech24-pulse {
          0% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0.7); }
          70% { transform: scale(1); box-shadow: 0 0 0 6px rgba(59, 130, 246, 0); }
          100% { transform: scale(0.95); box-shadow: 0 0 0 0 rgba(59, 130, 246, 0); }
        }
      `;
      document.head.appendChild(style);
    },

    update(message, state = "idle") {
      this.init();
      if (this.msgEl) this.msgEl.innerHTML = message;
      if (!this.statusIcon || !this.container) return;

      const stateMap = {
        idle: { color: "#94a3b8", animation: "none", border: "#334155" },
        running: { color: "#3b82f6", animation: "tech24-pulse 1.5s infinite", border: "#3b82f6" },
        success: { color: "#4ade80", animation: "none", border: "#4ade80" },
        error: { color: "#ef4444", animation: "none", border: "#ef4444" }
      };

      const style = stateMap[state] || stateMap.idle;
      this.statusIcon.style.background = style.color;
      this.statusIcon.style.animation = style.animation;
      this.container.style.borderColor = style.border;
    },

    showError(title, detail) {
      this.update(`<strong>${title}</strong><br><span style="font-size:12px; color:#cbd5e1;">${detail}</span>`, "error");
    }
  };

  // =========================================================================
  // 3. Network / Google Apps Script Bridge
  // =========================================================================
  const NetworkBridge = {
    gmGetJson(url) {
      return new Promise((resolve, reject) => {
        GM_xmlhttpRequest({
          method: "GET",
          url: url,
          timeout: 45000,
          onload: (res) => {
            const trimmed = (res.responseText || "").trim();

            if (trimmed.toLowerCase().startsWith("<!doctype html>")) {
              if (res.status === 404) return reject(new Error("Dashboard cannot find Google Apps Script. Check URL."));
              if (res.status === 429) return reject(new Error("Google rate limit hit."));
              return reject(new Error(`Google server error HTML (Status: ${res.status}).`));
            }

            if (res.status >= 200 && res.status < 300) {
              try {
                const data = JSON.parse(trimmed);
                if (data.error) return reject(new Error(`Sheet Error: ${data.error}`));
                resolve(data);
              } catch (e) {
                reject(new Error("Invalid JSON format received from Google."));
              }
            } else {
              reject(new Error(`Server refused connection (Error ${res.status}).`));
            }
          },
          onerror: () => reject(new Error("Network disconnect. Please check your internet connection.")),
          ontimeout: () => reject(new Error("Google Servers took too long to respond."))
        });
      });
    },

    get(action, extraParams = "") {
      const query = extraParams ? `${extraParams}&_=${Date.now()}` : `_=${Date.now()}`;
      return this.gmGetJson(`${CONFIG.BRIDGE_URL}?action=${action}&${query}`);
    }
  };

  // =========================================================================
  // 4. Date & Timestamp Parsing Helpers
  // =========================================================================
  const DateParser = {
    formatDateForSheet(rawDate) {
      if (!rawDate) return "";
      const clean = rawDate.replace(/[()]/g, "").trim();
      if (!clean) return "";

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
          // 3. Fallback: Native Date Object parsing
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
      const dayName = CONFIG.WEEKDAY_NAMES[dateObj.getDay()];
      const formattedMM = String(mm).padStart(2, "0");
      const formattedDD = String(dd).padStart(2, "0");
      const formattedYY = String(yyyy).slice(-2);

      return `${dayName} ${formattedMM}-${formattedDD}-${formattedYY}`;
    },

    splitTimestamp(text) {
      if (!text) return { date: null, time: null };
      const clean = text.replace(/[()]/g, "").replace(/^\(\vert{}\)$/g, "").trim();
      const lower = clean.toLowerCase();

      // Ignore empty/placeholder values for open or unclosed cases
      if (!clean || clean === "-" || lower.includes("pending") || lower.includes("open") || lower.includes("n/a")) {
        return { date: null, time: null };
      }

      const parts = clean.split(/\s+/);
      if (parts.length < 2) return { date: null, time: null };

      const dateStr = this.formatDateForSheet(parts[0]);
      const timeStr = parts.slice(1).join(" ");
      return { date: dateStr, time: timeStr };
    }
  };

  // =========================================================================
  // 5. DOM Extraction & Search Automation
  // =========================================================================
  const DashboardDOM = {
    getColumnIndices() {
      if (cachedCols) return cachedCols;
      const headerCells = Array.from(document.querySelectorAll("table thead th"));
      let found = { caseId: -1, start: -1, end: -1 };

      headerCells.forEach((th, idx) => {
        const label = th.textContent.toLowerCase().trim();
        if (label.includes("case id") || label.includes("caseid") || label.includes("case_id") || label === "case") {
          found.caseId = idx;
        } else if (label.includes("start") || label.includes("reg") || label.includes("opened") || label.includes("created")) {
          found.start = idx;
        } else if (label.includes("end") || label.includes("close") || label.includes("resolved") || label.includes("finished")) {
          found.end = idx;
        }
      });

      if (found.caseId === -1) found.caseId = 0;
      if (found.start === -1) found.start = 8;
      if (found.end === -1) found.end = 9;
      return (cachedCols = found);
    },

    getSearchInput() {
      return document.querySelector(CONFIG.SEARCH_SELECTOR) || document.querySelector('input[type="search"]');
    },

    setSearchValue(input, value) {
      if (!input) return;
      const nativeSetter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, "value")?.set;
      if (nativeSetter) {
        nativeSetter.call(input, value);
      } else {
        input.value = value;
      }
      input.dispatchEvent(new Event("input", { bubbles: true }));
      input.dispatchEvent(new Event("change", { bubbles: true }));
      input.dispatchEvent(new KeyboardEvent("keyup", { bubbles: true }));
    },

    findMatchingRow(caseId) {
      const cols = this.getColumnIndices();
      return Array.from(document.querySelectorAll("table tbody tr")).find((row) => {
        const cells = row.querySelectorAll("td");
        return cells.length && cells[cols.caseId]?.textContent.trim() === caseId;
      });
    },

    waitForFilteredRow(caseId) {
      return new Promise((resolve) => {
        const start = Date.now();
        const check = () => {
          const row = this.findMatchingRow(caseId);
          if (row) {
            resolve(row);
            return;
          }
          if (Date.now() - start > CONFIG.FILTER_WAIT_TIMEOUT_MS) return resolve(null);
          setTimeout(check, CONFIG.FILTER_POLL_MS);
        };
        check();
      });
    }
  };

  const delay = (ms) => new Promise((resolve) => setTimeout(resolve, ms));

  // =========================================================================
  // 6. Main Synchronization Engine
  // =========================================================================
  async function syncAll() {
    if (syncing) return;
    syncing = true;
    cachedCols = null;

    const stats = { inserted: 0, skipped: 0, leftBlank: 0 };

    try {
      SyncUI.update("Connecting to Google Sheets...", "running");

      const { rows } = await NetworkBridge.get("getPendingCases");

      if (!rows || rows.length === 0) {
        await NetworkBridge.get("clearTrigger", "summary={}");
        SyncUI.update("No empty rows found in sheet.", "idle");
        console.log("[Tech24 Sync] 📭 No pending cases found.");
        syncing = false;
        return;
      }

      console.log(`[Tech24 Sync] 🚀 Starting sync process... ${rows.length} rows ready to be filled.`);

      for (let i = 0; i < rows.length; i++) {
        const pending = rows[i];
        const input = DashboardDOM.getSearchInput();
        if (!input) throw new Error("Search bar not found on dashboard.");

        DashboardDOM.setSearchValue(input, pending.caseId);
        SyncUI.update(`Processing Case ${pending.caseId} (${i + 1}/${rows.length})...`, "running");

        const row = await DashboardDOM.waitForFilteredRow(pending.caseId);

        if (row) {
          const cols = DashboardDOM.getColumnIndices();
          const cells = row.querySelectorAll("td");
          const reg = DateParser.splitTimestamp(cells[cols.start]?.textContent.trim());
          const closed = DateParser.splitTimestamp(cells[cols.end]?.textContent.trim());

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

          let success = false;
          for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
            try {
              await NetworkBridge.get("submitResult", params.toString());
              console.log(`[Tech24 Sync] ✅ Updated case ${pending.caseId} (Start: ${reg.date || 'None'}, Closed: ${closed.date || 'Open'})`);
              success = true;
              break;
            } catch (submitErr) {
              console.warn(`[Tech24 Sync] ⚠️ Retry on case ${pending.caseId} (Attempt ${attempt}/${CONFIG.MAX_RETRIES}): ${submitErr.message}`);
              if (attempt < CONFIG.MAX_RETRIES) await delay(CONFIG.RETRY_DELAY_MS);
            }
          }

          if (!success) {
            console.error(`[Tech24 Sync] ❌ Failed to update case ${pending.caseId} after ${CONFIG.MAX_RETRIES} attempts.`);
          }

        } else {
          stats.skipped++;
          console.log(`[Tech24 Sync] ⏭️ Case not found on dashboard after 12s, skipped: ${pending.caseId}`);
        }
      }

      await NetworkBridge.get("clearTrigger", `summary=${JSON.stringify(stats)}`);
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

  // =========================================================================
  // 7. Event Handlers & Trigger Checks
  // =========================================================================
  async function checkAndRunSync(manualTrigger = false) {
    if (syncing) return;

    if (!manualTrigger && (Date.now() - lastCheckTime < CONFIG.CHECK_INTERVAL_MS)) return;
    lastCheckTime = Date.now();

    try {
      if (manualTrigger) SyncUI.update("Checking for trigger...", "running");
      const { shouldRun } = await NetworkBridge.get("checkTrigger");

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

  window.addEventListener("focus", () => checkAndRunSync(false));
  document.addEventListener("visibilitychange", () => {
    if (document.visibilityState === "visible") checkAndRunSync(false);
  });

  setTimeout(() => {
    SyncUI.init();
    checkAndRunSync(false);
  }, 1000);

  // Expose parser/dom helpers for debugging or internal extensions
  window.Tech24TimestampSync = {
    DateParser,
    DashboardDOM,
    SyncUI,
    checkAndRunSync
  };

})();
