// ==UserScript==
// @name         Tech24 SLA & Duration Columns
// @namespace    tech24-sla-duration
// @version      1.0
// @description  Calculates business-hour SLA durations for Tech24 dashboard.
// @match        https://tech24et.com/cases*
// @updateURL    https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-sla-duration.user.js
// @downloadURL  https://raw.githubusercontent.com/YoHighnessM/custom_scripts/main/tech24-sla-duration.user.js
// @grant        none
// ==/UserScript==

(function () {
    'use strict';

    // ---------------------------------------------------------------------
    // Config
    // ---------------------------------------------------------------------
    const ZONE_TARGET_HOURS = { '1': 7, '2': 8, '3': 14 };
    const BUSINESS_START_HOUR = 8;   // 8:00 AM
    const BUSINESS_END_HOUR = 17;    // 5:00 PM
    const BUSINESS_DAYS = [1, 2, 3, 4, 5, 6]; // Mon(1)..Sat(6), Sunday(0) excluded

    const CASE_ID_HEADER = 'Case ID';
    const START_HEADER = 'Start';
    const END_HEADER = 'End';
    const DISTRICT_HEADER = 'District';

    // ---------------------------------------------------------------------
    // Timestamp parsing/formatting
    // ---------------------------------------------------------------------
    function parseDashboardDate(str) {
        if (!str) return null;
        const trimmed = str.trim();
        if (trimmed === '-' || trimmed === '') return null;

        // Matches formats like "2026-08-14 01:51 PM" or "(2026-08-14 01:51 PM)"
        const match = trimmed.match(/^\(?\s*(\d{4})[/-](\d{2})[/-](\d{2})\s+(\d{1,2}):(\d{2})\s*(AM|PM)\s*\)?$/i);

        if (!match) {
            console.warn("[Tech24 SLA] Unrecognized date format:", str);
            return null;
        }

        let [, yyyy, mm, dd, hh, min, ampm] = match;
        let hour = parseInt(hh, 10);
        if (ampm.toUpperCase() === 'PM' && hour !== 12) hour += 12;
        if (ampm.toUpperCase() === 'AM' && hour === 12) hour = 0;

        return new Date(
            parseInt(yyyy, 10),
            parseInt(mm, 10) - 1,
            parseInt(dd, 10),
            hour,
            parseInt(min, 10),
            0
        );
    }

    function formatDateTime(date) {
        const pad = (n) => String(n).padStart(2, '0');
        let hour = date.getHours();
        const ampm = hour >= 12 ? 'PM' : 'AM';
        hour = hour % 12;
        if (hour === 0) hour = 12;
        return `${pad(date.getDate())}-${pad(date.getMonth() + 1)}-${date.getFullYear()} ${pad(
            hour
        )}:${pad(date.getMinutes())} ${ampm}`;
    }

    function formatDuration(totalMinutes) {
        const sign = totalMinutes < 0 ? '-' : '';
        const abs = Math.abs(Math.round(totalMinutes));
        const h = Math.floor(abs / 60);
        const m = abs % 60;
        return `${sign}${h}h ${m}m`;
    }

    // ---------------------------------------------------------------------
    // Business-hours-aware target time calculation
    // ---------------------------------------------------------------------
    function isBusinessDay(date) {
        return BUSINESS_DAYS.includes(date.getDay());
    }

    function nextBusinessDayStart(date) {
        const d = new Date(date);
        d.setDate(d.getDate() + 1);
        d.setHours(BUSINESS_START_HOUR, 0, 0, 0);
        while (!isBusinessDay(d)) {
            d.setDate(d.getDate() + 1);
        }
        return d;
    }

    function clampToBusinessWindow(date) {
        const d = new Date(date);
        if (isBusinessDay(d)) {
            const startOfDay = new Date(d);
            startOfDay.setHours(BUSINESS_START_HOUR, 0, 0, 0);
            const endOfDay = new Date(d);
            endOfDay.setHours(BUSINESS_END_HOUR, 0, 0, 0);

            if (d < startOfDay) return startOfDay;
            if (d >= endOfDay) return nextBusinessDayStart(d);
            return d;
        }
        return nextBusinessDayStart(d);
    }

    function addBusinessHours(start, hours) {
        let remainingMinutes = hours * 60;
        let cursor = clampToBusinessWindow(start);

        while (remainingMinutes > 0) {
            const endOfDay = new Date(cursor);
            endOfDay.setHours(BUSINESS_END_HOUR, 0, 0, 0);

            const minutesLeftToday = (endOfDay - cursor) / 60000;

            if (minutesLeftToday >= remainingMinutes) {
                cursor = new Date(cursor.getTime() + remainingMinutes * 60000);
                remainingMinutes = 0;
            } else {
                remainingMinutes -= minutesLeftToday;
                cursor = nextBusinessDayStart(cursor);
            }
        }
        return cursor;
    }

    function businessMinutesBetween(from, to) {
        let cursor = new Date(from);
        let total = 0;

        while (cursor < to) {
            if (!isBusinessDay(cursor)) {
                cursor = nextBusinessDayStart(cursor);
                continue;
            }
            const startOfDay = new Date(cursor);
            startOfDay.setHours(BUSINESS_START_HOUR, 0, 0, 0);
            const endOfDay = new Date(cursor);
            endOfDay.setHours(BUSINESS_END_HOUR, 0, 0, 0);

            if (cursor < startOfDay) {
                cursor = startOfDay;
                continue;
            }
            if (cursor >= endOfDay) {
                cursor = nextBusinessDayStart(cursor);
                continue;
            }

            const segmentEnd = to < endOfDay ? to : endOfDay;
            total += (segmentEnd - cursor) / 60000;
            cursor = segmentEnd >= endOfDay ? nextBusinessDayStart(segmentEnd) : segmentEnd;
        }
        return total;
    }

    function businessMinutesDiff(from, to) {
        if (from.getTime() === to.getTime()) return 0;
        if (from < to) return businessMinutesBetween(from, to);
        return -businessMinutesBetween(to, from);
    }

    // ---------------------------------------------------------------------
    // Zone persistence (per Case ID, via Tampermonkey storage)
    // ---------------------------------------------------------------------
    function getZone(caseId, defaultZone = '1') {
        return GM_getValue(`zone_${caseId}`, defaultZone);
    }

    function setZone(caseId, zone) {
        GM_setValue(`zone_${caseId}`, zone);
    }

    // ---------------------------------------------------------------------
    // Table discovery
    // ---------------------------------------------------------------------
    function findCaseTable() {
        const tables = document.querySelectorAll('table');
        for (const table of tables) {
            const headerCells = table.querySelectorAll('thead th');
            for (const th of headerCells) {
                if (th.textContent.trim().startsWith(CASE_ID_HEADER)) {
                    return table;
                }
            }
        }
        return null;
    }

    function getColumnIndex(table, label) {
        const headerCells = Array.from(table.querySelectorAll('thead th'));
        return headerCells.findIndex((th) => th.textContent.trim().startsWith(label));
    }

    // ---------------------------------------------------------------------
    // Header injection
    // ---------------------------------------------------------------------
    function ensureHeader(table) {
        const headerRow = table.querySelector('thead tr');
        if (!headerRow) return;

        if (!headerRow.querySelector('[data-sla-header]')) {
            const th = document.createElement('th');
            th.setAttribute('data-sla-header', 'true');
            th.textContent = 'SLA';
            headerRow.appendChild(th);
        }

        if (!headerRow.querySelector('[data-duration-header]')) {
            const th = document.createElement('th');
            th.setAttribute('data-duration-header', 'true');
            th.textContent = 'Duration';
            headerRow.appendChild(th);
        }
    }

    // ---------------------------------------------------------------------
    // Row rendering
    // ---------------------------------------------------------------------
    const activeTimers = new Map(); // caseId -> intervalId

    function renderRow(row, caseIdIdx, startIdx, endIdx, districtIdx) {
        const cells = row.querySelectorAll('td');
        if (cells.length <= Math.max(caseIdIdx, startIdx, endIdx)) return;

        const caseId = cells[caseIdIdx].textContent.trim();
        if (!caseId) return;

        const startText = cells[startIdx].textContent.trim();
        const endText = cells[endIdx].textContent.trim();

        let autoZone = '1';
        if (districtIdx !== -1 && cells.length > districtIdx) {
            const districtText = cells[districtIdx].textContent.trim().toLowerCase();
            if (districtText.includes('hawassa') || districtText.includes('wolayta')) {
                autoZone = '2';
            } else if (districtText.includes('shashemene')) {
                autoZone = '3';
            }
        }

        let slaCell = row.querySelector('[data-sla-cell]');
        if (!slaCell) {
            slaCell = document.createElement('td');
            slaCell.setAttribute('data-sla-cell', 'true');
            row.appendChild(slaCell);
        }

        let durationCell = row.querySelector('[data-duration-cell]');
        if (!durationCell) {
            durationCell = document.createElement('td');
            durationCell.setAttribute('data-duration-cell', 'true');
            row.appendChild(durationCell);
        }

        const zone = getZone(caseId, autoZone);
        const signature = `${caseId}|${startText}|${endText}|${zone}`;
        if (slaCell.dataset.slaSignature === signature) {
            return;
        }

        if (activeTimers.has(caseId)) {
            clearInterval(activeTimers.get(caseId));
            activeTimers.delete(caseId);
        }

        const startDate = parseDashboardDate(startText);
        const endDate = parseDashboardDate(endText);

        if (!startDate) {
            slaCell.textContent = '—';
            durationCell.textContent = '—';
            slaCell.dataset.slaSignature = signature;
            return;
        }

        const targetDate = addBusinessHours(startDate, ZONE_TARGET_HOURS[zone]);

        slaCell.innerHTML = '';
        slaCell.dataset.slaSignature = signature;

        const zoneSelect = document.createElement('select');
        zoneSelect.style.marginRight = '6px';
        zoneSelect.style.fontSize = '11px';
        ['1', '2', '3'].forEach((z) => {
            const opt = document.createElement('option');
            opt.value = z;
            opt.textContent = `Zone ${z}`;
            if (z === zone) opt.selected = true;
            zoneSelect.appendChild(opt);
        });
        zoneSelect.addEventListener('change', () => {
            setZone(caseId, zoneSelect.value);
            renderRow(row, caseIdIdx, startIdx, endIdx, districtIdx);
        });

        const textSpan = document.createElement('div');
        textSpan.style.fontSize = '12px';
        textSpan.style.lineHeight = '1.4';

        const targetLine = document.createElement('div');
        targetLine.textContent = `Target: ${formatDateTime(targetDate)}`;
        textSpan.appendChild(targetLine);

        const statusLine = document.createElement('div');
        textSpan.appendChild(statusLine);

        slaCell.appendChild(zoneSelect);
        slaCell.appendChild(textSpan);

        if (endDate) {
            // Closed case — static comparisons
            const diffMinutes = businessMinutesDiff(endDate, targetDate);
            if (diffMinutes >= 0) {
                statusLine.textContent = `Closed within SLA (${formatDuration(diffMinutes)} to spare)`;
                statusLine.style.color = '#16a34a';
            } else {
                statusLine.textContent = `Closed over SLA by ${formatDuration(-diffMinutes)}`;
                statusLine.style.color = '#dc2626';
            }

            // Duration restricted purely to business hours
            const bizDuration = businessMinutesDiff(startDate, endDate);
            durationCell.textContent = formatDuration(bizDuration);
            durationCell.style.fontSize = '12px';
        } else {
            // Open case — live countdown & live business duration
            const tick = () => {
                const now = new Date();
                const diffMinutes = businessMinutesDiff(now, targetDate);

                // Update SLA countdown
                if (diffMinutes >= 0) {
                    statusLine.textContent = `${formatDuration(diffMinutes)} remaining`;
                    statusLine.style.color = diffMinutes < 60 ? '#d97706' : '#16a34a';
                } else {
                    statusLine.textContent = `Overdue by ${formatDuration(-diffMinutes)}`;
                    statusLine.style.color = '#dc2626';
                }

                // Update Live Duration
                const currentDuration = businessMinutesDiff(startDate, now);
                durationCell.textContent = formatDuration(currentDuration);
                durationCell.style.fontSize = '12px';
                durationCell.style.color = '#6b7280';
            };
            tick();
            const intervalId = setInterval(tick, 1000);
            activeTimers.set(caseId, intervalId);
        }
    }

    // ---------------------------------------------------------------------
    // Main apply pass
    // ---------------------------------------------------------------------
    function apply() {
        const table = findCaseTable();
        if (!table) return;

        ensureHeader(table);

        const caseIdIdx = getColumnIndex(table, CASE_ID_HEADER);
        const startIdx = getColumnIndex(table, START_HEADER);
        const endIdx = getColumnIndex(table, END_HEADER);
        const districtIdx = getColumnIndex(table, DISTRICT_HEADER);

        if (caseIdIdx === -1 || startIdx === -1 || endIdx === -1) return;

        const rows = table.querySelectorAll('tbody tr');
        rows.forEach((row) => renderRow(row, caseIdIdx, startIdx, endIdx, districtIdx));
    }

    // ---------------------------------------------------------------------
    // Watch for dashboard re-renders
    // ---------------------------------------------------------------------
    let debounceId = null;
    function scheduleApply() {
        clearTimeout(debounceId);
        debounceId = setTimeout(apply, 150);
    }

    const OWN_SELECTOR = '[data-sla-cell],[data-sla-header],[data-duration-cell],[data-duration-header]';

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

    apply();
})();
