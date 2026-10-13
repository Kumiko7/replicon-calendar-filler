// ==UserScript==
// @name         Replicon Calendar Importer (v5 - Final Row/Selector Fix)
// @namespace    http://tampermonkey.net/
// @version      3.9
// @description  Robustly finds the newly created row and project selector link before proceeding.
// @author       Gemini
// @match        https://*.replicon.com/*/my/timesheet/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        unsafeWindow
// @connect      localhost
// @run-at       document-idle
// ==/UserScript==

(function() {
    'use strict';

    const SCRIPT_ID = 'replicon-calendar-importer';
    const BUTTON_ID = 'fetch-calendar-button';

    const SELECTORS = {
        buttonContainer: '.customTimesheetHeader',
        dayBody: (dateStr) => `tbody[aria-label="${dateStr}"]`,
        workRow: 'tr.work',
        breakRow: 'tr.break',
        timeOffRow: 'tr[class*="timeoff" i], tr[class*="time-off" i]',
        commentTextarea: 'td.comments textarea[aria-label="Comments"]',
        workInTimeInput: 'input[aria-label="In"]',
        workOutTimeInput: 'input[aria-label="Out"]',
        breakStartTimeInput: 'input[aria-label="In"]',
        breakEndTimeInput: 'input[aria-label="Out"]',
        // --- CRITICAL FIX: A more stable selector for the project link ---
        projectDropdownButton: "a.multiLevelSelector.TaskSelectorSearchFixedWidth[aria-label='Type to search']",
        projectDropdownPanel: 'body > div.divDropdownContent[is-wormhole][style*="display: block"]',
        projectSearchInput: 'input[placeholder="Type to search all assigned"]',
        projectFirstResult: '.searchAllListContainer ul.divDropdownListTable a.taskSelectorSearchAllRowItem',
    };

    // Pause between events so Replicon can finish saving the previous row
    const DELAY_BETWEEN_EVENTS_MS = 300;
    // How long to wait for project search results before giving up
    const PROJECT_SEARCH_TIMEOUT_MS = 10000;
    // If results stop changing for this long without an exact match, assume the search has finished
    const PROJECT_RESULTS_SETTLE_MS = 1500;

    let isRunning = false;

    const sleep = (ms) => new Promise(resolve => setTimeout(resolve, ms));

    function waitForElement(selector, parent = document, timeout = 10000) {
        return new Promise((resolve, reject) => {
            const intervalTime = 100;
            let timeWaited = 0;
            const interval = setInterval(() => {
                const element = parent.querySelector(selector);
                if (element) {
                    clearInterval(interval);
                    resolve(element);
                } else {
                    timeWaited += intervalTime;
                    if (timeWaited >= timeout) {
                        clearInterval(interval);
                        reject(new Error(`Element "${selector}" not found after ${timeout}ms`));
                    }
                }
            }, intervalTime);
        });
    }

    function waitForNewRow(dayBody, rowSelector, initialCount, timeout) {
        return new Promise((resolve, reject) => {
            const interval = setInterval(() => {
                const rows = dayBody.querySelectorAll(rowSelector);
                if (rows.length > initialCount) {
                    clearInterval(interval);
                    clearTimeout(timer);
                    resolve(rows[initialCount]);
                }
            }, 100);
            const timer = setTimeout(() => { clearInterval(interval); reject(new Error('New row did not appear.')); }, timeout);
        });
    }

    function parseDateTimeAsLocal(dateString) {
        const match = dateString.match(/^(\d{4})-(\d{2})-(\d{2})T(\d{2}):(\d{2}):(\d{2})/);
        if (!match) return new Date(NaN);
        return new Date(
            parseInt(match[1]), parseInt(match[2]) - 1, parseInt(match[3]),
            parseInt(match[4]), parseInt(match[5]), parseInt(match[6])
        );
    }

    function formatTime(date) {
        return `${String(date.getHours()).padStart(2, '0')}:${String(date.getMinutes()).padStart(2, '0')}`;
    }

    function formatDateForAriaLabel(date) {
        return date.toLocaleDateString('en-US', { month: 'short', day: 'numeric', year: 'numeric' });
    }

    function describeEvent(event) {
        const start = parseDateTimeAsLocal(event.start);
        const end = parseDateTimeAsLocal(event.end);
        const subject = event.subject || '(blank)';
        return `${formatDateForAriaLabel(start)} ${formatTime(start)}-${formatTime(end)} "${subject}"`;
    }

    function clickActionLink(parentElement, linkText) {
        const link = Array.from(parentElement.querySelectorAll('a.action')).find(a => a.textContent.trim() === linkText);
        if (link) {
            link.click();
            return true;
        }
        throw new Error(`Could not find "${linkText}" button.`);
    }

    function setInputValue(input, value, events = ['change', 'blur']) {
        input.value = value;
        for (const type of events) {
            input.dispatchEvent(new Event(type, { bubbles: true }));
        }
    }

    function getDayBody(event) {
        const eventDateStr = formatDateForAriaLabel(parseDateTimeAsLocal(event.start));
        const dayBody = document.querySelector(SELECTORS.dayBody(eventDateStr));
        if (!dayBody) throw new Error(`Day "${eventDateStr}" is not on this timesheet.`);
        return dayBody;
    }

    async function addBreakToReplicon(event) {
        const startDate = parseDateTimeAsLocal(event.start);
        const endDate = parseDateTimeAsLocal(event.end);
        const dayBody = getDayBody(event);

        const initialBreakCount = dayBody.querySelectorAll(SELECTORS.breakRow).length;
        clickActionLink(dayBody, '+ Break Time');
        const newBreakRow = await waitForNewRow(dayBody, SELECTORS.breakRow, initialBreakCount, 5000);

        const startTimeInput = newBreakRow.querySelector(SELECTORS.breakStartTimeInput);
        const endTimeInput = newBreakRow.querySelector(SELECTORS.breakEndTimeInput);
        if (!startTimeInput || !endTimeInput) throw new Error('Could not find all required fields for the new break entry.');

        setInputValue(startTimeInput, formatTime(startDate));
        setInputValue(endTimeInput, formatTime(endDate));
    }

    function getProjectResults() {
        const panel = document.querySelector(SELECTORS.projectDropdownPanel);
        return panel ? Array.from(panel.querySelectorAll(SELECTORS.projectFirstResult)) : [];
    }

    function findExactProjectMatch(results, project) {
        return results.find(result => {
            const projectElement = result.querySelector('div[data-id="project"]');
            return projectElement && projectElement.textContent.trim() === project;
        });
    }

    // Returns true if the project was selected, false if no exact match was found.
    async function clickAndFillProject(workRow, project) {
        console.log(`[Project] Starting selection for: "${project}"`);

        const projectButton = await waitForElement(SELECTORS.projectDropdownButton, workRow, 5000);
        projectButton.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
        projectButton.dispatchEvent(new Event('focus', { bubbles: true }));
        projectButton.dispatchEvent(new MouseEvent('click', { bubbles: true }));

        const searchInput = await waitForElement(SELECTORS.projectSearchInput, workRow);
        setInputValue(searchInput, project, ['input', 'keyup']);

        // Poll for results instead of waiting a fixed time: click as soon as the exact
        // project shows up, or give up once the results have stopped changing.
        const startedAt = Date.now();
        let lastSnapshot = null;
        let lastChangeAt = Date.now();
        while (Date.now() - startedAt < PROJECT_SEARCH_TIMEOUT_MS) {
            const results = getProjectResults();
            const exactMatch = findExactProjectMatch(results, project);
            if (exactMatch) {
                exactMatch.click();
                projectButton.dispatchEvent(new Event('blur', { bubbles: true }));
                console.log(`[Project] SUCCESS: Selected "${project}".`);
                return true;
            }

            const snapshot = results.map(r => r.textContent.trim()).join('|');
            if (snapshot !== lastSnapshot) {
                lastSnapshot = snapshot;
                lastChangeAt = Date.now();
            } else if (results.length > 0 && Date.now() - lastChangeAt >= PROJECT_RESULTS_SETTLE_MS) {
                break;
            }
            await sleep(150);
        }

        console.warn(`[Project] No exact match for "${project}". Leaving project blank.`);
        document.body.click(); // Close the dropdown without selecting anything
        return false;
    }

    // Returns a warning message if the row was added but needs attention, otherwise null.
    async function addWorkEventToReplicon(event) {
        const startDate = parseDateTimeAsLocal(event.start);
        const endDate = parseDateTimeAsLocal(event.end);
        const dayBody = getDayBody(event);

        const initialWorkRowCount = dayBody.querySelectorAll(SELECTORS.workRow).length;
        clickActionLink(dayBody, '+ Work Time');
        const newWorkRow = await waitForNewRow(dayBody, SELECTORS.workRow, initialWorkRowCount, 10000);

        const inTimeInput = newWorkRow.querySelector(SELECTORS.workInTimeInput);
        const outTimeInput = newWorkRow.querySelector(SELECTORS.workOutTimeInput);
        if (!inTimeInput || !outTimeInput) throw new Error('Could not find In/Out time inputs.');

        setInputValue(inTimeInput, formatTime(startDate));
        setInputValue(outTimeInput, formatTime(endDate));

        const commentTextarea = newWorkRow.querySelector(SELECTORS.commentTextarea);
        if (commentTextarea) {
            setInputValue(commentTextarea, event.subject, ['input', 'change', 'blur']);
        }

        if (event.project && event.project.trim() !== "") {
            let selected = false;
            try {
                selected = await clickAndFillProject(newWorkRow, event.project);
            } catch (error) {
                console.error(`[Project] FAILED to select project "${event.project}". Error: ${error.message}`);
                document.body.click(); // Attempt to close any open dialogs
            }
            if (!selected) {
                newWorkRow.style.outline = '2px solid #dc3545';
                return `project "${event.project}" not found, left blank (row outlined in red)`;
            }
        }
        return null;
    }

    function retrieveTimesheetDates() {
        const model = unsafeWindow.model;
        if (typeof model === 'object' && model && model.details && model.details.dateRange) {
            const dateObjToApiFormat = (dateObj) => `${dateObj.year}-${String(dateObj.month).padStart(2, '0')}-${String(dateObj.day).padStart(2, '0')}`;
            return {
                startDate: dateObjToApiFormat(model.details.dateRange.startDate),
                endDate: dateObjToApiFormat(model.details.dateRange.endDate)
            };
        }
        return null;
    }

    function countExistingEntries() {
        const inputs = document.querySelectorAll(
            `${SELECTORS.workRow} ${SELECTORS.workInTimeInput}, ${SELECTORS.breakRow} ${SELECTORS.breakStartTimeInput}`
        );
        return Array.from(inputs).filter(input => input.value.trim() !== '').length;
    }

    function resetButton(button, text = 'Fetch Calendar Data') {
        button.textContent = text;
        button.disabled = false;
    }

    // A day has time off (e.g. Annual Leave) booked in Replicon if it has a time-off row.
    function dayHasTimeOff(dayBody) {
        if (dayBody.querySelector(SELECTORS.timeOffRow)) return true;
        // Fallback: time-off rows show the leave balance, e.g. "72.00 hours taken"
        return Array.from(dayBody.querySelectorAll('tr')).some(row =>
            !row.matches(`${SELECTORS.workRow}, ${SELECTORS.breakRow}`) && /hours taken/i.test(row.textContent));
    }

    async function populateTimesheet(allEvents, button) {
        const warnings = [];
        const failures = [];

        // Skip events on days that already have time off booked in Replicon
        const timeOffDays = new Set();
        const calendarData = allEvents.filter(event => {
            const dayBody = document.querySelector(SELECTORS.dayBody(formatDateForAriaLabel(parseDateTimeAsLocal(event.start))));
            if (dayBody && dayHasTimeOff(dayBody)) {
                timeOffDays.add(dayBody.getAttribute('aria-label'));
                return false;
            }
            return true;
        });
        const skippedCount = allEvents.length - calendarData.length;

        for (const [index, event] of calendarData.entries()) {
            button.textContent = `Populating ${index + 1}/${calendarData.length}...`;
            try {
                if (event.subject === "Break Time") {
                    await addBreakToReplicon(event);
                } else {
                    const warning = await addWorkEventToReplicon(event);
                    if (warning) warnings.push(`${describeEvent(event)}: ${warning}`);
                }
            } catch (error) {
                console.error(`${SCRIPT_ID} Failed on ${describeEvent(event)}:`, error);
                failures.push(`${describeEvent(event)}: ${error.message}`);
            }
            await sleep(DELAY_BETWEEN_EVENTS_MS);
        }

        const okCount = calendarData.length - warnings.length - failures.length;
        let summary = `Finished: ${okCount} of ${calendarData.length} entries added successfully.`;
        if (skippedCount) summary += `\n\nSkipped ${skippedCount} events on days with time off: ${[...timeOffDays].join(', ')}`;
        if (warnings.length) summary += `\n\nNeeds attention (${warnings.length}):\n- ${warnings.join('\n- ')}`;
        if (failures.length) summary += `\n\nFailed, not added (${failures.length}):\n- ${failures.join('\n- ')}`;
        alert(summary);

        resetButton(button, (warnings.length || failures.length) ? 'Done (with issues)' : 'Done!');
    }

    async function fetchAndFillData(button) {
        if (isRunning) return;
        isRunning = true;
        try {
            await runImport(button);
        } finally {
            isRunning = false;
        }
    }

    async function runImport(button) {
        const dates = retrieveTimesheetDates();
        if (!dates) {
            handleError(new Error('Could not read the timesheet dates. The page may not have finished loading - try again in a moment.'), button);
            return;
        }

        const existing = countExistingEntries();
        if (existing > 0 && !confirm(
            `This timesheet already has ${existing} time ${existing === 1 ? 'entry' : 'entries'}.\n\n` +
            `Importing will add calendar events on top of them, which may create duplicates.\n\nContinue anyway?`)) {
            return;
        }

        button.textContent = 'Fetching Calendar...';
        button.disabled = true;

        const apiUrl = `http://localhost:8000/calendar?from=${dates.startDate}&to=${dates.endDate}`;
        try {
            const response = await new Promise((resolve, reject) => GM_xmlhttpRequest({
                method: 'GET',
                url: apiUrl,
                onload: resolve,
                onerror: () => reject(new Error('Failed to fetch from localhost. Is your local server running?')),
            }));
            if (response.status < 200 || response.status >= 300) {
                throw new Error(`Server responded with status: ${response.status} ${response.statusText}`);
            }
            const calendarData = JSON.parse(response.responseText);
            if (calendarData.length === 0) {
                alert('No calendar events found for this timesheet period.');
                resetButton(button);
                return;
            }
            await populateTimesheet(calendarData, button);
        } catch (err) {
            handleError(err, button);
        }
    }

    function handleError(error, button) {
        console.error(`${SCRIPT_ID} Error:`, error.message);
        alert(`An error occurred: ${error.message}`);
        button.textContent = 'Error!';
        button.disabled = true;
        setTimeout(() => resetButton(button), 5000);
    }

    function injectButton(container) {
        if (document.getElementById(BUTTON_ID)) return;
        const fetchButton = document.createElement('button');
        fetchButton.textContent = 'Fetch Calendar Data';
        fetchButton.id = BUTTON_ID;
        fetchButton.addEventListener('click', () => fetchAndFillData(fetchButton));
        container.appendChild(fetchButton);
    }

    // --- Initialization ---
    GM_addStyle(`
        #${BUTTON_ID} {
            padding: 8px 16px; font-size: 14px; font-weight: bold; color: #fff;
            background-color: #007bff; border: none; border-radius: 4px;
            cursor: pointer; margin-left: 20px;
            transition: background-color 0.2s;
            float: right;
        }
        #${BUTTON_ID}:hover { background-color: #0056b3; }
        #${BUTTON_ID}:disabled { background-color: #cccccc; cursor: not-allowed; }
    `);

    function ensureButton() {
        const container = document.querySelector(SELECTORS.buttonContainer);
        if (container && !document.getElementById(BUTTON_ID)) {
            injectButton(container);
        }
    }

    // Keep watching: Replicon redraws the header when switching timesheet periods,
    // which removes the button, so it has to be re-added each time.
    new MutationObserver(ensureButton).observe(document.documentElement, { childList: true, subtree: true });
    ensureButton();

})();
