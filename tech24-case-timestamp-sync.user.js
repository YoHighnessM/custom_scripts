// ==UserScript==
// @name         Tech24 Case Timestamp Sync (Ultra-Fast Edition)
// @namespace    tech24-case-sla-sync
// @version      5.0
// @description  Direct API extraction, dynamic logging, flexible date formatting, open-case handling, and Technicians Activity Tracker.
// @match        https://tech24et.com/cases*
// @match        https://dashboard.tech24et.com/cases*
// @match        https://dashboard.tech24et.com/*
// @updateURL    https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-case-timestamp-sync.user.js
// @downloadURL  https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-case-timestamp-sync.user.js
// @grant        GM_xmlhttpRequest
// @connect      script.google.com
// @connect      script.googleusercontent.com
// @connect      docs.google.com
// @connect      api.dashboard.tech24et.com
// @connect      *
// ==/UserScript==

(function () {
  "use strict";

  // =========================================================================
  // 1. Configuration & Constants
  // =========================================================================
  const CONFIG = Object.freeze({
    BRIDGE_URL: "https://script.google.com/macros/s/AKfycbywn6bKxqdObHkgB4x7AgqyLLSCBYH2Mk5rSUfL_MndKiPfgeF8BRPd8o-TCGGOIbhI-g/exec",
    SPREADSHEET_ID: "1_JUmZ_LoU1wrjoLGnh7HvE3AXRFU5eGkFg6fS5o_BsM",
    TAB_NAME: "Technicians Activity Summary by District",
    API_BASE_URL: "https://api.dashboard.tech24et.com/api/callentries",
    SEARCH_SELECTOR: 'input[placeholder="Search Cases..."]',
    FILTER_WAIT_TIMEOUT_MS: 12000,
    FILTER_POLL_MS: 50,
    CHECK_INTERVAL_MS: 2000,
    API_RATE_LIMIT_DELAY_MS: 500,
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
      const clean = String(rawDate).replace(/[()]/g, "").trim();
      if (!clean) return "";

      let yyyy, mm, dd;

      // 1. Check ISO format with T (e.g. 2026-10-29T11:56:04.000000Z)
      if (clean.includes("T")) {
        const parsed = new Date(clean);
        if (!isNaN(parsed.getTime())) {
          yyyy = String(parsed.getFullYear());
          mm = String(parsed.getMonth() + 1);
          dd = String(parsed.getDate());
        }
      }

      if (!yyyy) {
        // 2. Check YYYY-MM-DD or YYYY/MM/DD
        let match = clean.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})/);
        if (match) {
          [, yyyy, mm, dd] = match;
        } else {
          // 3. Check MM-DD-YYYY or MM/DD/YYYY
          match = clean.match(/^(\d{1,2})[/-](\d{1,2})[/-](\d{4})/);
          if (match) {
            [, mm, dd, yyyy] = match;
          } else {
            // 4. Fallback: Native Date Object parsing
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
      }

      const dateObj = new Date(parseInt(yyyy, 10), parseInt(mm, 10) - 1, parseInt(dd, 10));
      const dayName = CONFIG.WEEKDAY_NAMES[dateObj.getDay()];
      const formattedMM = String(mm).padStart(2, "0");
      const formattedDD = String(dd).padStart(2, "0");
      const formattedYY = String(yyyy).slice(-2);

      return `${dayName} ${formattedMM}-${formattedDD}-${formattedYY}`;
    },

    formatTimeFromDate(dateObj) {
      if (!dateObj || isNaN(dateObj.getTime())) return null;
      let hours = dateObj.getHours();
      const minutes = String(dateObj.getMinutes()).padStart(2, "0");
      const ampm = hours >= 12 ? "PM" : "AM";
      hours = hours % 12;
      if (hours === 0) hours = 12;
      const formattedHours = String(hours).padStart(2, "0");
      return `${formattedHours}:${minutes} ${ampm}`;
    },

    splitTimestamp(text) {
      if (!text) return { date: null, time: null };
      const clean = String(text).replace(/[()]/g, "").replace(/^\(\vert{}\)$/g, "").trim();
      const lower = clean.toLowerCase();

      // Ignore empty/placeholder values for open or unclosed cases
      if (!clean || clean === "-" || lower.includes("pending") || lower.includes("open") || lower.includes("n/a")) {
        return { date: null, time: null };
      }

      // Check if text is ISO timestamp format (e.g. 2026-10-29T11:56:04.000000Z)
      if (clean.includes("T")) {
        const d = new Date(clean);
        if (!isNaN(d.getTime())) {
          return {
            date: this.formatDateForSheet(clean),
            time: this.formatTimeFromDate(d)
          };
        }
      }

      const parts = clean.split(/\s+/);
      if (parts.length < 2) return { date: null, time: null };

      const dateStr = this.formatDateForSheet(parts[0]);
      const timeStr = parts.slice(1).join(" ");
      return { date: dateStr, time: timeStr };
    }
  };

  // =========================================================================
  // 5. API Extraction Service
  // =========================================================================
  const ApiExtractor = {
    extractCaseFields(result, requestedCaseId) {
      if (!result || !Array.isArray(result.data)) {
        throw new Error("Unexpected API response: data array missing");
      }

      const requestedId = String(requestedCaseId).trim();

      const record = result.data.find(
        (item) => String(item.callentry_id).trim() === requestedId
      );

      if (!record) {
        return {
          caseId: requestedId,
          found: false,
          assignedAt: "",
          completedAt: "",
          createdAt: "",
          updatedAt: "",
          repeatedDates: [],
          error: "Case not found"
        };
      }

      return {
        caseId: String(record.callentry_id).trim(),
        found: true,
        assignedAt: record.assigned_at ?? "",
        completedAt: record.completed_at ?? "",
        createdAt: record.created_at ?? "",
        updatedAt: record.updated_at ?? "",
        repeatedDates: Array.isArray(record.repeated_callentries_dates)
          ? record.repeated_callentries_dates
          : [],
        error: ""
      };
    },

    async fetchCaseFromApi(caseId) {
      const params = new URLSearchParams({
        limit: "20",
        search: String(caseId).trim(),
        callstatus: "",
        approvalstatus: "",
        start_date: "",
        end_date: "",
        active: "",
        bank: "",
        branch: "",
        district: ""
      });

      const url = `${CONFIG.API_BASE_URL}?${params.toString()}`;

      // 1. First attempt: normal fetch with credentials
      try {
        if (typeof fetch === "function") {
          const response = await fetch(url, {
            method: "GET",
            credentials: "include",
            headers: {
              Accept: "application/json"
            }
          });

          if (response.ok) {
            const result = await response.json();
            return this.extractCaseFields(result, caseId);
          } else if (response.status === 401 || response.status === 403) {
            throw new Error(`Authentication required (HTTP ${response.status}). Please log in to Tech24.`);
          } else if (response.status === 429) {
            throw new Error(`Rate limit hit (HTTP 429).`);
          }
        }
      } catch (err) {
        // If it was an auth error, re-throw immediately
        if (err.message && err.message.includes("Authentication required")) {
          throw err;
        }
        // Otherwise fallback to GM_xmlhttpRequest
      }

      // 2. Fallback: GM_xmlhttpRequest
      return await this.fetchCaseFromApiGM(url, caseId);
    },

    fetchCaseFromApiGM(url, caseId) {
      return new Promise((resolve, reject) => {
        if (typeof GM_xmlhttpRequest !== "function") {
          return reject(new Error("GM_xmlhttpRequest not available."));
        }

        GM_xmlhttpRequest({
          method: "GET",
          url: url,
          headers: {
            Accept: "application/json"
          },
          timeout: 20000,
          onload: (response) => {
            if (response.status === 401 || response.status === 403) {
              return reject(new Error(`Authentication required (HTTP ${response.status}). Please log in.`));
            }
            if (response.status === 429) {
              return reject(new Error(`Rate limit hit (HTTP 429).`));
            }
            if (response.status < 200 || response.status >= 300) {
              return reject(new Error(`API returned HTTP ${response.status}`));
            }

            try {
              const data = JSON.parse(response.responseText);
              resolve(this.extractCaseFields(data, caseId));
            } catch (err) {
              reject(new Error("API response was not valid JSON"));
            }
          },
          onerror: () => reject(new Error("Network request to API failed")),
          ontimeout: () => reject(new Error("API request timed out"))
        });
      });
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

      console.log(`[Tech24 Sync] 🚀 Starting API sync process... ${rows.length} rows ready to be filled.`);

      for (let i = 0; i < rows.length; i++) {
        const pending = rows[i];
        SyncUI.update(`Processing Case ${pending.caseId} (${i + 1}/${rows.length})...`, "running");

        let caseData = null;
        for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
          try {
            caseData = await ApiExtractor.fetchCaseFromApi(pending.caseId);
            break;
          } catch (apiErr) {
            console.warn(`[Tech24 Sync] ⚠️ API attempt ${attempt}/${CONFIG.MAX_RETRIES} for case ${pending.caseId}: ${apiErr.message}`);
            if (apiErr.message && apiErr.message.includes("Authentication required")) {
              throw apiErr; // Don't retry auth errors, prompt user to log in
            }
            if (attempt < CONFIG.MAX_RETRIES) {
              await delay(CONFIG.RETRY_DELAY_MS * attempt);
            }
          }
        }

        if (caseData && caseData.found) {
          // Parse start timestamp: prefer created_at, fallback to assigned_at
          const rawStart = caseData.createdAt || caseData.assignedAt;
          const reg = DateParser.splitTimestamp(rawStart);
          const closed = DateParser.splitTimestamp(caseData.completedAt);

          const hasDataToInsert = Boolean(reg.date || closed.date);
          if (!hasDataToInsert) {
            stats.leftBlank++;
          } else {
            stats.inserted++;
          }

          const params = new URLSearchParams({
            row: String(pending.row),
            caseId: String(pending.caseId),
            regDate: reg.date || "",
            regTime: reg.time || "",
            closedDate: closed.date || "",
            closedTime: closed.time || ""
          });

          let submitSuccess = false;
          for (let attempt = 1; attempt <= CONFIG.MAX_RETRIES; attempt++) {
            try {
              await NetworkBridge.get("submitResult", params.toString());
              console.log(`[Tech24 Sync] ✅ Updated case ${pending.caseId} (Start: ${reg.date || 'None'} ${reg.time || ''}, Closed: ${closed.date || 'Open'} ${closed.time || ''})`);
              submitSuccess = true;
              break;
            } catch (submitErr) {
              console.warn(`[Tech24 Sync] ⚠️ Retry on submitting case ${pending.caseId} (Attempt ${attempt}/${CONFIG.MAX_RETRIES}): ${submitErr.message}`);
              if (attempt < CONFIG.MAX_RETRIES) await delay(CONFIG.RETRY_DELAY_MS);
            }
          }

          if (!submitSuccess) {
            console.error(`[Tech24 Sync] ❌ Failed to update case ${pending.caseId} after ${CONFIG.MAX_RETRIES} attempts.`);
          }

        } else {
          stats.skipped++;
          console.log(`[Tech24 Sync] ⏭️ Case not found or error for ${pending.caseId}: ${caseData ? caseData.error : 'No response'}`);
        }

        // Pacing delay between API requests to respect rate limits
        if (CONFIG.API_RATE_LIMIT_DELAY_MS > 0 && i < rows.length - 1) {
          await delay(CONFIG.API_RATE_LIMIT_DELAY_MS);
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

  // =========================================================================
  // 8. Technicians Activity Summary Tracker (Modal & Info Icon)
  // =========================================================================
  const ActivityTracker = {
    TABLE_NAMES: ["Central Addis", "Wolayta", "Hawassa", "Shashemene"],
    cachedData: null,
    isLoading: false,
    modalEl: null,
    iconEl: null,

    init() {
      this.injectStyles();
      this.ensureInfoIcon();
      this.observeHeader();
    },

    injectStyles() {
      if (document.getElementById("tech24-activity-tracker-style")) return;
      const style = document.createElement("style");
      style.id = "tech24-activity-tracker-style";
      style.textContent = `
        .tech24-info-icon-btn {
          display: inline-flex;
          align-items: center;
          justify-content: center;
          width: 22px;
          height: 22px;
          border-radius: 50%;
          background: #eff6ff;
          color: #2563eb;
          border: 1px solid #bfdbfe;
          cursor: pointer;
          transition: all 0.2s ease;
          margin-right: 6px;
          padding: 0;
          outline: none;
          vertical-align: middle;
        }
        .tech24-info-icon-btn:hover {
          background: #2563eb;
          color: #ffffff;
          border-color: #1d4ed8;
          transform: scale(1.08);
          box-shadow: 0 2px 5px rgba(37,99,235,0.25);
        }
        .tech24-activity-modal-overlay {
          position: fixed;
          top: 0;
          left: 0;
          width: 100vw;
          height: 100vh;
          background: rgba(15, 23, 42, 0.65);
          backdrop-filter: blur(4px);
          z-index: 100000;
          display: flex;
          align-items: center;
          justify-content: center;
          padding: 20px;
          box-sizing: border-box;
          opacity: 0;
          pointer-events: none;
          transition: opacity 0.25s ease;
        }
        .tech24-activity-modal-overlay.open {
          opacity: 1;
          pointer-events: auto;
        }
        .tech24-activity-modal-content {
          background: #ffffff;
          color: #0f172a;
          border-radius: 12px;
          max-width: 950px;
          width: 100%;
          max-height: 85vh;
          overflow: hidden;
          display: flex;
          flex-direction: column;
          box-shadow: 0 20px 25px -5px rgba(0,0,0,0.2), 0 10px 10px -5px rgba(0,0,0,0.1);
          border: 1px solid #e2e8f0;
          font-family: system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", Roboto, sans-serif;
        }
        .tech24-modal-header {
          display: flex;
          align-items: center;
          justify-content: space-between;
          padding: 16px 24px;
          border-bottom: 1px solid #e2e8f0;
          background: #f8fafc;
        }
        .tech24-modal-title {
          font-size: 17px;
          font-weight: 700;
          color: #1e293b;
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .tech24-modal-title span.tag {
          font-size: 11px;
          font-weight: 600;
          color: #2563eb;
          background: #dbeafe;
          padding: 2px 8px;
          border-radius: 9999px;
        }
        .tech24-modal-actions {
          display: flex;
          align-items: center;
          gap: 10px;
        }
        .tech24-modal-refresh-btn, .tech24-modal-close-btn {
          border: none;
          background: transparent;
          border-radius: 6px;
          padding: 6px;
          cursor: pointer;
          color: #64748b;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          transition: all 0.2s;
        }
        .tech24-modal-refresh-btn:hover {
          background: #e2e8f0;
          color: #2563eb;
        }
        .tech24-modal-close-btn:hover {
          background: #fee2e2;
          color: #ef4444;
        }
        .tech24-spin {
          animation: tech24-spin-kf 0.8s linear infinite;
        }
        @keyframes tech24-spin-kf {
          from { transform: rotate(0deg); }
          to { transform: rotate(360deg); }
        }
        .tech24-modal-body {
          padding: 20px 24px;
          overflow-y: auto;
          flex-grow: 1;
        }
        .tech24-district-section {
          margin-bottom: 24px;
        }
        .tech24-district-title {
          font-size: 14px;
          font-weight: 700;
          color: #334155;
          margin-bottom: 8px;
          display: flex;
          align-items: center;
          gap: 8px;
        }
        .tech24-district-title::before {
          content: "";
          display: inline-block;
          width: 8px;
          height: 8px;
          border-radius: 50%;
          background: #3b82f6;
        }
        .tech24-table-custom {
          width: 100%;
          border-collapse: separate;
          border-spacing: 0;
          border: 1px solid #e2e8f0;
          border-radius: 8px;
          overflow: hidden;
          font-size: 13px;
        }
        .tech24-table-custom th {
          background: #f1f5f9;
          color: #475569;
          font-weight: 600;
          text-align: left;
          padding: 8px 12px;
          border-bottom: 1px solid #e2e8f0;
          border-right: 1px solid #f1f5f9;
        }
        .tech24-table-custom td {
          padding: 8px 12px;
          border-bottom: 1px solid #f1f5f9;
          color: #1e293b;
        }
        .tech24-table-custom tr:last-child td {
          border-bottom: none;
        }
        .tech24-table-custom tr:hover td {
          background: #f8fafc;
        }
        .tech24-empty-msg {
          color: #94a3b8;
          font-style: italic;
          padding: 12px;
          text-align: center;
        }
      `;
      document.head.appendChild(style);
    },

    ensureInfoIcon() {
      // Find header row in dashboard case table
      const headerRow = document.querySelector("table thead tr");
      if (!headerRow) return;

      // Find the SLA header element either by attribute or text content
      let slaTh = headerRow.querySelector("[data-sla-header]");
      if (!slaTh) {
        const headerCells = Array.from(headerRow.children);
        slaTh = headerCells.find((th) => {
          const txt = (th.textContent || "").trim();
          return txt === "SLA" || (txt.startsWith("SLA") && !txt.includes("Status"));
        });
      }
      if (!slaTh) return;

      // Check if button is already injected
      if (slaTh.querySelector("#tech24-info-icon-btn")) return;

      const btn = document.createElement("button");
      btn.id = "tech24-info-icon-btn";
      btn.type = "button";
      btn.className = "tech24-info-icon-btn";
      btn.title = "View Technicians Activity Summary by District";
      btn.innerHTML = `
        <svg xmlns="http://www.w3.org/2000/svg" width="13" height="13" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2.5" stroke-linecap="round" stroke-linejoin="round">
          <circle cx="12" cy="12" r="10"></circle>
          <line x1="12" y1="16" x2="12" y2="12"></line>
          <line x1="12" y1="8" x2="12.01" y2="8"></line>
        </svg>
      `;

      btn.addEventListener("click", (e) => {
        e.preventDefault();
        e.stopPropagation();
        this.openModal();
      });

      // Insert inside the SLA header element before the SLA label text
      const flexContainer = slaTh.querySelector("div.flex") || slaTh.querySelector("div") || slaTh;
      if (flexContainer.firstChild) {
        flexContainer.insertBefore(btn, flexContainer.firstChild);
      } else {
        flexContainer.appendChild(btn);
      }
    },

    observeHeader() {
      const observer = new MutationObserver(() => {
        this.ensureInfoIcon();
      });
      observer.observe(document.body, { childList: true, subtree: true });
    },

    buildModalDOM() {
      if (this.modalEl) return this.modalEl;

      const overlay = document.createElement("div");
      overlay.className = "tech24-activity-modal-overlay";
      overlay.id = "tech24-activity-tracker-modal";

      overlay.innerHTML = `
        <div class="tech24-activity-modal-content" onclick="event.stopPropagation()">
          <div class="tech24-modal-header">
            <div class="tech24-modal-title">
              Technicians Activity Summary
              <span class="tag">By District</span>
            </div>
            <div class="tech24-modal-actions">
              <button class="tech24-modal-refresh-btn" id="tech24-tracker-refresh-btn" title="Refresh from Google Sheet">
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                  <path d="M21.5 2v6h-6M21.34 15.57a10 10 0 1 1-.57-8.38l5.67-5.67"/>
                </svg>
              </button>
              <button class="tech24-modal-close-btn" id="tech24-tracker-close-btn" title="Close">
                <svg xmlns="http://www.w3.org/2000/svg" width="18" height="18" viewBox="0 0 24 24" fill="none" stroke="currentColor" stroke-width="2" stroke-linecap="round" stroke-linejoin="round">
                  <line x1="18" y1="6" x2="6" y2="18"></line>
                  <line x1="6" y1="6" x2="18" y2="18"></line>
                </svg>
              </button>
            </div>
          </div>
          <div class="tech24-modal-body" id="tech24-tracker-modal-body">
            <div style="text-align:center; padding: 40px; color:#64748b;">Loading activity tracker data...</div>
          </div>
        </div>
      `;

      overlay.onclick = () => this.closeModal();

      const closeBtn = overlay.querySelector("#tech24-tracker-close-btn");
      closeBtn.onclick = () => this.closeModal();

      const refreshBtn = overlay.querySelector("#tech24-tracker-refresh-btn");
      refreshBtn.onclick = () => this.refreshData();

      document.body.appendChild(overlay);

      window.addEventListener("keydown", (e) => {
        if (e.key === "Escape" && overlay.classList.contains("open")) {
          this.closeModal();
        }
      });

      this.modalEl = overlay;
      return overlay;
    },

    openModal() {
      const modal = this.buildModalDOM();
      modal.classList.add("open");
      if (!this.cachedData) {
        this.refreshData();
      } else {
        this.renderBody(this.cachedData);
      }
    },

    closeModal() {
      if (this.modalEl) {
        this.modalEl.classList.remove("open");
      }
    },

    async refreshData() {
      if (this.isLoading) return;
      this.isLoading = true;

      const refreshBtn = document.getElementById("tech24-tracker-refresh-btn");
      if (refreshBtn) {
        const svg = refreshBtn.querySelector("svg");
        if (svg) svg.classList.add("tech24-spin");
      }

      const bodyEl = document.getElementById("tech24-tracker-modal-body");
      if (bodyEl && !this.cachedData) {
        bodyEl.innerHTML = `<div style="text-align:center; padding: 40px; color:#64748b;">Fetching latest data from Google Sheet...</div>`;
      }

      try {
        const data = await this.fetchData();
        this.cachedData = data;
        this.renderBody(data);
      } catch (err) {
        console.error("[ActivityTracker Error]:", err);
        if (bodyEl) {
          bodyEl.innerHTML = `
            <div style="padding:20px; background:#fef2f2; border:1px solid #fecaca; border-radius:8px; color:#991b1b; text-align:center;">
              <strong>Failed to load Activity Tracker data</strong><br>
              <span style="font-size:12px; color:#b91c1c;">${err.message}</span>
            </div>
          `;
        }
      } finally {
        this.isLoading = false;
        if (refreshBtn) {
          const svg = refreshBtn.querySelector("svg");
          if (svg) svg.classList.remove("tech24-spin");
        }
      }
    },

    async fetchData() {
      // 1. Try Google Apps Script bridge endpoint
      try {
        const res = await NetworkBridge.get("getTechniciansActivitySummary");
        if (res && res.data && typeof res.data === "object") {
          return res.data;
        }
      } catch (bridgeErr) {
        console.warn("[ActivityTracker] Bridge fetch failed, attempting CSV fallback:", bridgeErr.message);
      }

      // 2. Direct Fallback: fetch via Google Sheets CSV export or visualization API
      const exportUrl = `https://docs.google.com/spreadsheets/d/${CONFIG.SPREADSHEET_ID}/gviz/tq?tqx=out:csv&sheet=${encodeURIComponent(CONFIG.TAB_NAME)}`;
      return await this.fetchFromCsvUrl(exportUrl);
    },

    fetchFromCsvUrl(url) {
      return new Promise((resolve, reject) => {
        if (typeof GM_xmlhttpRequest !== "function") {
          return reject(new Error("GM_xmlhttpRequest not available."));
        }

        GM_xmlhttpRequest({
          method: "GET",
          url: url,
          timeout: 25000,
          onload: (response) => {
            if (response.status >= 200 && response.status < 300) {
              const text = response.responseText || "";
              if (text.includes("<!DOCTYPE html>") && text.includes("accounts.google.com")) {
                return reject(new Error("Access denied. Please ensure the Google Sheet is shared or accessed via the bridge."));
              }
              const parsed = this.parseCsvData(text);
              resolve(parsed);
            } else {
              reject(new Error(`Failed to download sheet CSV (HTTP ${response.status})`));
            }
          },
          onerror: () => reject(new Error("Network error fetching sheet data")),
          ontimeout: () => reject(new Error("Timeout fetching sheet data"))
        });
      });
    },

    parseCsvData(csvText) {
      const lines = csvText.split(/\r?\n/).map((line) => {
        // Simple CSV splitter handling quotes
        const result = [];
        let cur = "";
        let inQuotes = false;
        for (let i = 0; i < line.length; i++) {
          const char = line[i];
          if (char === '"' && line[i + 1] === '"') {
            cur += '"';
            i++;
          } else if (char === '"') {
            inQuotes = !inQuotes;
          } else if (char === "," && !inQuotes) {
            result.push(cur.trim());
            cur = "";
          } else {
            cur += char;
          }
        }
        result.push(cur.trim());
        return result;
      });

      const tables = {};
      this.TABLE_NAMES.forEach((name) => {
        tables[name] = [];
      });

      this.TABLE_NAMES.forEach((name) => {
        const norm = name.toLowerCase().replace(/[^a-z0-9]/g, "");
        for (let r = 0; r < lines.length; r++) {
          for (let c = 0; c < lines[r].length; c++) {
            const cell = (lines[r][c] || "").toLowerCase().replace(/[^a-z0-9]/g, "");
            if (cell === norm) {
              // Locate header row below
              for (let subR = r + 1; subR < Math.min(lines.length, r + 5); subR++) {
                const rowSlice = lines[subR].slice(c, c + 7);
                if (rowSlice.some((v) => v.toLowerCase().includes("engineer"))) {
                  const headers = rowSlice.map((h) => h.toLowerCase());
                  const engIdx = headers.findIndex((h) => h.includes("engineer"));
                  const physIdx = headers.findIndex((h) => h.includes("physical"));
                  const phoneIdx = headers.findIndex((h) => h.includes("phone"));
                  const totalCaseIdx = headers.findIndex((h) => h.includes("total case"));
                  const pmIdx = headers.findIndex((h) => h === "pm" || h.includes("pm"));
                  const totalActIdx = headers.findIndex((h) => h.includes("total activity"));

                  for (let dataR = subR + 1; dataR < lines.length; dataR++) {
                    const dataRow = lines[dataR].slice(c, c + 7);
                    const eng = (dataRow[engIdx !== -1 ? engIdx : 0] || "").trim();
                    if (!eng || eng.toLowerCase() === "total" || eng.toLowerCase() === "engineers") break;

                    const physical = (dataRow[physIdx !== -1 ? physIdx : 1] || "0").trim();
                    const phone = (dataRow[phoneIdx !== -1 ? phoneIdx : 2] || "0").trim();
                    const totalCase = (dataRow[totalCaseIdx !== -1 ? totalCaseIdx : 3] || "0").trim();
                    const pm = (dataRow[pmIdx !== -1 ? pmIdx : 4] || "0").trim();
                    const totalActivity = (dataRow[totalActIdx !== -1 ? totalActIdx : 5] || "0").trim();

                    tables[name].push({
                      engineer: eng,
                      physicalAddress: physical,
                      phoneSupport: phone,
                      totalCase,
                      pm,
                      totalActivity,
                      totalCaseDisplay: `${totalCase} (${physical} | ${phone})`
                    });
                  }
                  break;
                }
              }
            }
          }
        }
      });

      return tables;
    },

    renderBody(data) {
      const bodyEl = document.getElementById("tech24-tracker-modal-body");
      if (!bodyEl) return;

      bodyEl.innerHTML = "";

      let hasAnyRows = false;
      this.TABLE_NAMES.forEach((district) => {
        const rows = data && data[district] ? data[district] : [];
        if (rows.length > 0) hasAnyRows = true;

        const section = document.createElement("div");
        section.className = "tech24-district-section";

        const title = document.createElement("div");
        title.className = "tech24-district-title";
        title.textContent = district;
        section.appendChild(title);

        if (rows.length === 0) {
          const empty = document.createElement("div");
          empty.className = "tech24-empty-msg";
          empty.textContent = "No technician activity records found for this district.";
          section.appendChild(empty);
        } else {
          const table = document.createElement("table");
          table.className = "tech24-table-custom";

          // Table Header: Engineers, Total Case (Physical Address | Phone Support), PM, Total Activity
          table.innerHTML = `
            <thead>
              <tr>
                <th style="width: 28%;">Engineers</th>
                <th style="width: 36%;">Total Case (Physical Address | Phone Support)</th>
                <th style="width: 16%; text-align: center;">PM</th>
                <th style="width: 20%; text-align: center;">Total Activity</th>
              </tr>
            </thead>
            <tbody>
              ${rows.map((r) => `
                <tr>
                  <td style="font-weight: 500;">${r.engineer || "—"}</td>
                  <td>${r.totalCaseDisplay || `${r.totalCase || 0} (${r.physicalAddress || 0} | ${r.phoneSupport || 0})`}</td>
                  <td style="text-align: center; color: #475569;">${r.pm || 0}</td>
                  <td style="text-align: center; font-weight: 600; color: #2563eb;">${r.totalActivity || 0}</td>
                </tr>
              `).join("")}
            </tbody>
          `;
          section.appendChild(table);
        }

        bodyEl.appendChild(section);
      });

      if (!hasAnyRows) {
        bodyEl.innerHTML = `
          <div style="text-align:center; padding:30px; color:#64748b;">
            <p style="font-size:14px; margin-bottom:8px;">No activity data currently available.</p>
            <p style="font-size:12px; color:#94a3b8;">Click the refresh icon above once data is entered in Google Sheets.</p>
          </div>
        `;
      }
    }
  };

  // Initialize immediately and also after brief timeouts to handle dynamic SPA route transitions
  ActivityTracker.init();
  SyncUI.init();

  setTimeout(() => {
    SyncUI.init();
    ActivityTracker.init();
    checkAndRunSync(false);
  }, 1000);

  // Expose parser/dom/api helpers for debugging or internal extensions
  window.Tech24TimestampSync = {
    DateParser,
    DashboardDOM,
    ApiExtractor,
    SyncUI,
    ActivityTracker,
    checkAndRunSync
  };

})();
