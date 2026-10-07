// ==UserScript==
// @name         Tech24 SLA & Duration Columns
// @namespace    tech24-sla-duration
// @version      1.2
// @description  Calculates business-hour SLA durations for Tech24 dashboard.
// @match        https://tech24et.com/cases*
// @updateURL    https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-sla-duration.user.js
// @downloadURL  https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-sla-duration.user.js
// @grant        none
// ==/UserScript==

(function () {
  "use strict";

  // =========================================================================
  // 1. Configuration
  // =========================================================================
  const CONFIG = Object.freeze({
    ZONE_TARGET_HOURS: Object.freeze({ "1": 7, "2": 8, "3": 14 }),
    BUSINESS_START_HOUR: 8,  // 8:00 AM
    BUSINESS_END_HOUR: 17,   // 5:00 PM
    BUSINESS_DAYS: Object.freeze([1, 2, 3, 4, 5, 6]), // Mon(1)..Sat(6), Sunday(0) excluded

    CASE_ID_HEADER: "case id",
    START_HEADER: "start",
    END_HEADER: "end",
    DISTRICT_HEADER: "district"
  });

  // Active timers for open-case live countdowns (caseId -> intervalId)
  const activeTimers = new Map();

  // =========================================================================
  // 2. Local Storage Manager
  // =========================================================================
  const StorageManager = {
    getZone(caseId, defaultZone = "1") {
      try {
        if (typeof GM_getValue === "function") {
          return GM_getValue(`zone_${caseId}`, defaultZone);
        }
        return localStorage.getItem(`tech24_zone_${caseId}`) || defaultZone;
      } catch (e) {
        return defaultZone;
      }
    },

    setZone(caseId, zone) {
      try {
        if (typeof GM_setValue === "function") {
          GM_setValue(`zone_${caseId}`, zone);
        } else {
          localStorage.setItem(`tech24_zone_${caseId}`, zone);
        }
      } catch (e) {
        console.warn("[Tech24 SLA] Storage set error:", e);
      }
    }
  };

  // =========================================================================
  // 3. Timestamp Parsing & Formatting
  // =========================================================================
  const DateParser = {
    parseDashboardDate(str) {
      if (!str) return null;
      const trimmed = str.replace(/[()]/g, "").trim();
      if (!trimmed || trimmed === "-") return null;

      // Matches formats like "2026-08-14 01:51 PM", "2026/08/14 01:51 PM"
      const match = trimmed.match(/^(\d{4})[/-](\d{1,2})[/-](\d{1,2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)?$/i);

      if (!match) {
        // Fallback parser for standard dates
        const fallback = new Date(trimmed);
        return isNaN(fallback.getTime()) ? null : fallback;
      }

      let [, yyyy, mm, dd, hh, min, ampm] = match;
      let hour = parseInt(hh, 10);
      if (ampm) {
        const isPM = ampm.toUpperCase() === "PM";
        if (isPM && hour !== 12) hour += 12;
        if (!isPM && hour === 12) hour = 0;
      }

      return new Date(
        parseInt(yyyy, 10),
        parseInt(mm, 10) - 1,
        parseInt(dd, 10),
        hour,
        parseInt(min, 10),
        0
      );
    },

    formatDateTime(date) {
      const pad = (n) => String(n).padStart(2, "0");
      let hour = date.getHours();
      const ampm = hour >= 12 ? "PM" : "AM";
      hour = hour % 12;
      if (hour === 0) hour = 12;

      return `${pad(date.getDate())}-${pad(date.getMonth() + 1)}-${date.getFullYear()} ${pad(
        hour
      )}:${pad(date.getMinutes())} ${ampm}`;
    },

    formatDuration(totalMinutes) {
      const sign = totalMinutes < 0 ? "-" : "";
      const abs = Math.abs(Math.round(totalMinutes));
      const h = Math.floor(abs / 60);
      const m = abs % 60;
      return `${sign}${h}h ${m}m`;
    }
  };

  // =========================================================================
  // 4. Business Hours Calendar Calculations
  // =========================================================================
  const BusinessCalendar = {
    isBusinessDay(date) {
      return CONFIG.BUSINESS_DAYS.includes(date.getDay());
    },

    nextBusinessDayStart(date) {
      const d = new Date(date);
      d.setDate(d.getDate() + 1);
      d.setHours(CONFIG.BUSINESS_START_HOUR, 0, 0, 0);
      while (!this.isBusinessDay(d)) {
        d.setDate(d.getDate() + 1);
      }
      return d;
    },

    clampToBusinessWindow(date) {
      const d = new Date(date);
      if (this.isBusinessDay(d)) {
        const startOfDay = new Date(d);
        startOfDay.setHours(CONFIG.BUSINESS_START_HOUR, 0, 0, 0);
        const endOfDay = new Date(d);
        endOfDay.setHours(CONFIG.BUSINESS_END_HOUR, 0, 0, 0);

        if (d < startOfDay) return startOfDay;
        if (d >= endOfDay) return this.nextBusinessDayStart(d);
        return d;
      }
      return this.nextBusinessDayStart(d);
    },

    addBusinessHours(start, hours) {
      let remainingMinutes = hours * 60;
      let cursor = this.clampToBusinessWindow(start);

      while (remainingMinutes > 0) {
        const endOfDay = new Date(cursor);
        endOfDay.setHours(CONFIG.BUSINESS_END_HOUR, 0, 0, 0);

        const minutesLeftToday = (endOfDay - cursor) / 60000;

        if (minutesLeftToday >= remainingMinutes) {
          cursor = new Date(cursor.getTime() + remainingMinutes * 60000);
          remainingMinutes = 0;
        } else {
          remainingMinutes -= minutesLeftToday;
          cursor = this.nextBusinessDayStart(cursor);
        }
      }
      return cursor;
    },

    businessMinutesBetween(from, to) {
      let cursor = new Date(from);
      let total = 0;

      while (cursor < to) {
        if (!this.isBusinessDay(cursor)) {
          cursor = this.nextBusinessDayStart(cursor);
          continue;
        }
        const startOfDay = new Date(cursor);
        startOfDay.setHours(CONFIG.BUSINESS_START_HOUR, 0, 0, 0);
        const endOfDay = new Date(cursor);
        endOfDay.setHours(CONFIG.BUSINESS_END_HOUR, 0, 0, 0);

        if (cursor < startOfDay) {
          cursor = startOfDay;
          continue;
        }
        if (cursor >= endOfDay) {
          cursor = this.nextBusinessDayStart(cursor);
          continue;
        }

        const segmentEnd = to < endOfDay ? to : endOfDay;
        total += (segmentEnd - cursor) / 60000;
        cursor = segmentEnd >= endOfDay ? this.nextBusinessDayStart(segmentEnd) : segmentEnd;
      }
      return total;
    },

    businessMinutesDiff(from, to) {
      if (from.getTime() === to.getTime()) return 0;
      if (from < to) return this.businessMinutesBetween(from, to);
      return -this.businessMinutesBetween(to, from);
    }
  };

  // =========================================================================
  // 5. Table Extraction & Injection Engine
  // =========================================================================
  const TableRenderer = {
    findCaseTable() {
      const tables = document.querySelectorAll("table");
      for (const table of tables) {
        const headerCells = table.querySelectorAll("thead th");
        for (const th of headerCells) {
          const text = th.textContent.toLowerCase().trim();
          if (text.includes(CONFIG.CASE_ID_HEADER) || text === "caseid" || text === "case") {
            return table;
          }
        }
      }
      return null;
    },

    getColumnIndex(table, headerKey) {
      const headerCells = Array.from(table.querySelectorAll("thead th"));
      return headerCells.findIndex((th) => {
        const text = th.textContent.toLowerCase().trim();
        if (headerKey === CONFIG.CASE_ID_HEADER) {
          return text.includes("case id") || text.includes("caseid") || text === "case";
        }
        if (headerKey === CONFIG.START_HEADER) {
          return text.includes("start") || text.includes("reg") || text.includes("opened");
        }
        if (headerKey === CONFIG.END_HEADER) {
          return text.includes("end") || text.includes("close") || text.includes("resolved");
        }
        if (headerKey === CONFIG.DISTRICT_HEADER) {
          return text.includes("district") || text.includes("zone") || text.includes("region");
        }
        return false;
      });
    },

    ensureHeaders(table) {
      const headerRow = table.querySelector("thead tr");
      if (!headerRow) return;

      if (!headerRow.querySelector("[data-sla-header]")) {
        const th = document.createElement("th");
        th.className = "py-2 px-1 whitespace-nowrap";
        th.setAttribute("data-sla-header", "true");
        th.innerHTML = '<div class="flex gap-2 items-center">SLA</div>';
        headerRow.appendChild(th);
      }

      if (!headerRow.querySelector("[data-duration-header]")) {
        const th = document.createElement("th");
        th.className = "py-2 px-1 whitespace-nowrap";
        th.setAttribute("data-duration-header", "true");
        th.innerHTML = '<div class="flex gap-2 items-center">Duration</div>';
        headerRow.appendChild(th);
      }
    },

    renderRow(row, caseIdIdx, startIdx, endIdx, districtIdx) {
      const cells = row.querySelectorAll("td");
      if (cells.length <= Math.max(caseIdIdx, startIdx, endIdx)) return;

      const caseId = cells[caseIdIdx].textContent.trim();
      if (!caseId) return;

      const startText = cells[startIdx].textContent.trim();
      const endText = cells[endIdx].textContent.trim();

      let autoZone = "1";
      if (districtIdx !== -1 && cells.length > districtIdx) {
        const districtText = cells[districtIdx].textContent.trim().toLowerCase();
        if (districtText.includes("hawassa") || districtText.includes("wolaita")) {
          autoZone = "2";
        } else if (districtText.includes("shashemene")) {
          autoZone = "3";
        }
      }

      let slaCell = row.querySelector("[data-sla-cell]");
      if (!slaCell) {
        slaCell = document.createElement("td");
        slaCell.className = "py-2 px-1 whitespace-nowrap";
        slaCell.setAttribute("data-sla-cell", "true");
        row.appendChild(slaCell);
      }

      let durationCell = row.querySelector("[data-duration-cell]");
      if (!durationCell) {
        durationCell = document.createElement("td");
        durationCell.className = "py-2 px-1 whitespace-nowrap";
        durationCell.setAttribute("data-duration-cell", "true");
        row.appendChild(durationCell);
      }

      const zone = StorageManager.getZone(caseId, autoZone);
      const signature = `${caseId}|${startText}|${endText}|${zone}`;
      if (slaCell.dataset.slaSignature === signature) {
        return;
      }

      // Clear existing active countdown timer for this case if present
      if (activeTimers.has(caseId)) {
        clearInterval(activeTimers.get(caseId));
        activeTimers.delete(caseId);
      }

      const startDate = DateParser.parseDashboardDate(startText);
      const endDate = DateParser.parseDashboardDate(endText);

      if (!startDate) {
        slaCell.textContent = "—";
        durationCell.textContent = "—";
        slaCell.dataset.slaSignature = signature;
        return;
      }

      const targetHours = CONFIG.ZONE_TARGET_HOURS[zone] || 7;
      const targetDate = BusinessCalendar.addBusinessHours(startDate, targetHours);

      slaCell.innerHTML = "";
      slaCell.dataset.slaSignature = signature;

      // Render Zone selector dropdown with clean light design matching dashboard UI
      const zoneSelect = document.createElement("select");
      zoneSelect.style.cssText = "margin-right: 6px; font-size: 11px; font-weight: 500; padding: 2px 6px; border-radius: 6px; border: 1px solid #cbd5e1; background-color: #f8fafc; color: #1e293b; cursor: pointer; outline: none; transition: border-color 0.2s;";
      ["1", "2", "3"].forEach((z) => {
        const opt = document.createElement("option");
        opt.value = z;
        opt.textContent = `Zone ${z}`;
        if (z === zone) opt.selected = true;
        zoneSelect.appendChild(opt);
      });

      zoneSelect.addEventListener("change", (e) => {
        e.stopPropagation();
        StorageManager.setZone(caseId, zoneSelect.value);
        this.renderRow(row, caseIdIdx, startIdx, endIdx, districtIdx);
      });

      const textSpan = document.createElement("div");
      textSpan.style.cssText = "display: inline-block; font-size: 12px; line-height: 1.4; vertical-align: middle;";

      const targetLine = document.createElement("div");
      targetLine.style.color = "#475569";
      targetLine.textContent = `Target: ${DateParser.formatDateTime(targetDate)}`;
      textSpan.appendChild(targetLine);

      const statusLine = document.createElement("div");
      statusLine.style.fontWeight = "500";
      textSpan.appendChild(statusLine);

      slaCell.appendChild(zoneSelect);
      slaCell.appendChild(textSpan);

      if (endDate) {
        // Closed case — static calculation
        const diffMinutes = BusinessCalendar.businessMinutesDiff(endDate, targetDate);
        if (diffMinutes >= 0) {
          statusLine.textContent = `Within SLA (${DateParser.formatDuration(diffMinutes)} spare)`;
          statusLine.style.color = "#16a34a";
        } else {
          statusLine.textContent = `Over SLA by ${DateParser.formatDuration(-diffMinutes)}`;
          statusLine.style.color = "#dc2626";
        }

        const bizDuration = BusinessCalendar.businessMinutesDiff(startDate, endDate);
        durationCell.textContent = DateParser.formatDuration(bizDuration);
        durationCell.style.fontSize = "12px";
      } else {
        // Open case — live countdown & live business duration update
        const tick = () => {
          const now = new Date();
          const diffMinutes = BusinessCalendar.businessMinutesDiff(now, targetDate);

          if (diffMinutes >= 0) {
            statusLine.textContent = `${DateParser.formatDuration(diffMinutes)} remaining`;
            statusLine.style.color = diffMinutes < 60 ? "#d97706" : "#16a34a";
          } else {
            statusLine.textContent = `Overdue by ${DateParser.formatDuration(-diffMinutes)}`;
            statusLine.style.color = "#dc2626";
          }

          const currentDuration = BusinessCalendar.businessMinutesDiff(startDate, now);
          durationCell.textContent = DateParser.formatDuration(currentDuration);
          durationCell.style.fontSize = "12px";
          durationCell.style.color = "#64748b";
        };

        tick();
        const intervalId = setInterval(tick, 1000);
        activeTimers.set(caseId, intervalId);
      }
    },

    apply() {
      const table = this.findCaseTable();
      if (!table) return;

      this.ensureHeaders(table);

      const caseIdIdx = this.getColumnIndex(table, CONFIG.CASE_ID_HEADER);
      const startIdx = this.getColumnIndex(table, CONFIG.START_HEADER);
      const endIdx = this.getColumnIndex(table, CONFIG.END_HEADER);
      const districtIdx = this.getColumnIndex(table, CONFIG.DISTRICT_HEADER);

      if (caseIdIdx === -1 || startIdx === -1 || endIdx === -1) return;

      const rows = table.querySelectorAll("tbody tr");
      rows.forEach((row) => this.renderRow(row, caseIdIdx, startIdx, endIdx, districtIdx));
    }
  };

  // =========================================================================
  // 6. Observer & Execution Management
  // =========================================================================
  let debounceId = null;
  function scheduleApply() {
    clearTimeout(debounceId);
    debounceId = setTimeout(() => TableRenderer.apply(), 150);
  }

  const OWN_SELECTOR = "[data-sla-cell],[data-sla-header],[data-duration-cell],[data-duration-header]";

  function isOwnTarget(node) {
    const el = node.nodeType === Node.ELEMENT_NODE ? node : node.parentElement;
    if (!el || !el.closest) return false;
    return !!(el.closest(OWN_SELECTOR) || el.matches?.(OWN_SELECTOR));
  }

  function handleMutations(mutations) {
    const allOwn = mutations.every((m) => isOwnTarget(m.target));
    if (!allOwn) {
      scheduleApply();
    }
  }

  const observer = new MutationObserver(handleMutations);
  observer.observe(document.body, { childList: true, subtree: true });

  TableRenderer.apply();

  // Expose module for testing and debugging
  window.Tech24SLADuration = {
    DateParser,
    BusinessCalendar,
    StorageManager,
    TableRenderer,
    CONFIG
  };

})();
