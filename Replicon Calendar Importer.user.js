// ==UserScript==
// @name         Replicon Calendar Importer (v5 - Final Row/Selector Fix)
// @namespace    http://tampermonkey.net/
// @version      3.7
// @description  Robustly finds the newly created row and project selector link before proceeding.
// @author       Gemini
// @match        https://*.replicon.com/*/my/timesheet/*
// @grant        GM_xmlhttpRequest
// @grant        GM_addStyle
// @grant        unsafeWindow
// @run-at       document-start
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

    function clickActionLink(parentElement, linkText) {
        const link = Array.from(parentElement.querySelectorAll('a.action')).find(a => a.textContent.trim() === linkText);
        if (link) {
            link.click();
            return true;
        }
        throw new Error(`Could not find "${linkText}" button.`);
    }

    async function addBreakToReplicon(event) {
        const startDate = parseDateTimeAsLocal(event.start);
        const endDate = parseDateTimeAsLocal(event.end);
        const eventDateStr = formatDateForAriaLabel(startDate);
        const dayBody = document.querySelector(SELECTORS.dayBody(eventDateStr));
        if (!dayBody) return;

        const initialBreakCount = dayBody.querySelectorAll(SELECTORS.breakRow).length;
        clickActionLink(dayBody, '+ Break Time');

        await new Promise((resolve, reject) => {
            const interval = setInterval(() => {
                if (dayBody.querySelectorAll(SELECTORS.breakRow).length > initialBreakCount) {
                    clearInterval(interval); resolve();
                }
            }, 100);
            setTimeout(() => { clearInterval(interval); reject(new Error('New break row did not appear.')); }, 5000);
        });

        const newBreakRow = dayBody.querySelectorAll(SELECTORS.breakRow)[initialBreakCount];
        const startTimeInput = newBreakRow.querySelector(SELECTORS.breakStartTimeInput);
        const endTimeInput = newBreakRow.querySelector(SELECTORS.breakEndTimeInput);
        if (!startTimeInput || !endTimeInput) throw new Error('Could not find all required fields for the new break entry.');

        startTimeInput.value = formatTime(startDate);
        startTimeInput.dispatchEvent(new Event('change', { bubbles: true }));
        startTimeInput.dispatchEvent(new Event('blur', { bubbles: true }));
        endTimeInput.value = formatTime(endDate);
        endTimeInput.dispatchEvent(new Event('change', { bubbles: true }));
        endTimeInput.dispatchEvent(new Event('blur', { bubbles: true }));
    }


     async function clickAndFillProject(workRow, project) {
        try {
            console.log(`[Project] Starting selection for: "${project}"`);

            await new Promise(resolve => setTimeout(resolve, 500));

            const projectButton = await waitForElement(SELECTORS.projectDropdownButton, workRow, 5000);
            console.log("[Project] Found project selector button inside the new row.");

            projectButton.dispatchEvent(new MouseEvent('mousedown', { bubbles: true }));
            projectButton.dispatchEvent(new Event('focus', { bubbles: true }));
            projectButton.dispatchEvent(new MouseEvent('click', { bubbles: true }));
            console.log("[Project] Dispatched mousedown, focus, and click events.");

            const searchInput = await waitForElement(SELECTORS.projectSearchInput, workRow);
            console.log("[Project] Found search input inside panel.");

            searchInput.value = project;
            searchInput.dispatchEvent(new Event('input', { bubbles: true }));
            searchInput.dispatchEvent(new Event('keyup', { bubbles: true }));
            console.log(`[Project] Typed "${project}" into search input.`);

            await new Promise(resolve => setTimeout(resolve, 5000));

            const panel = await waitForElement(SELECTORS.projectDropdownPanel, document.body, 2000);
            console.log("[Project] Found visible dropdown panel in document body.");

            // MODIFICATION START
            // Get all search results instead of waiting for just the first one.
            const allResults = panel.querySelectorAll(SELECTORS.projectFirstResult);
            console.log(`[Project] Found ${allResults.length} total results.`);

            if (allResults.length === 0) {
                throw new Error('No project results found in the dropdown.');
            }

            let exactMatchElement = null;

            // Loop through all results to find an exact match for the project name.
            for (const result of allResults) {
                const projectElement = result.querySelector('div[data-id="project"]');
                if (projectElement && projectElement.textContent.trim() === project) {
                    console.log(`[Project] Found an exact match for "${project}".`);
                    exactMatchElement = result;
                    break; // Found the first exact match, so stop looking.
                }
            }

            // If an exact match was found, click it.
            if (exactMatchElement) {
                exactMatchElement.click();
                console.log(`[Project] SUCCESS: Clicked the exact match.`);
            } else {
                // Otherwise, fall back to clicking the very first result in the list.
                console.warn(`[Project] No exact match for "${project}" found. Falling back to the first result.`);
                const firstResult = allResults[0];
                firstResult.click();
                console.log(`[Project] SUCCESS (Fallback): Clicked the first overall result.`);
            }
            // MODIFICATION END

        } catch (error) {
            console.error(`[Project] FAILED to select project "${project}". Error: ${error.message}`);
            document.body.click(); // Attempt to close any open dialogs
        }
    }


    async function addWorkEventToReplicon(event) {
        const startDate = parseDateTimeAsLocal(event.start);
        const endDate = parseDateTimeAsLocal(event.end);
        const eventDateStr = formatDateForAriaLabel(startDate);
        const dayBody = document.querySelector(SELECTORS.dayBody(eventDateStr));
        if (!dayBody) return;

        const initialWorkRowCount = dayBody.querySelectorAll(SELECTORS.workRow).length;
        clickActionLink(dayBody, '+ Work Time');

        await new Promise((resolve, reject) => {
            const interval = setInterval(() => {
                if (dayBody.querySelectorAll(SELECTORS.workRow).length > initialWorkRowCount) {
                    clearInterval(interval); resolve();
                }
            }, 100);
            setTimeout(() => { clearInterval(interval); reject(new Error('New work row did not appear.')); }, 10000);
        });

        const newWorkRow = dayBody.querySelectorAll(SELECTORS.workRow)[initialWorkRowCount];
        console.log("Successfully found new work row.");

        const inTimeInput = newWorkRow.querySelector(SELECTORS.workInTimeInput);
        const outTimeInput = newWorkRow.querySelector(SELECTORS.workOutTimeInput);
        if (!inTimeInput || !outTimeInput) throw new Error('Could not find In/Out time inputs.');

        inTimeInput.value = formatTime(startDate);
        inTimeInput.dispatchEvent(new Event('change', { bubbles: true }));
        inTimeInput.dispatchEvent(new Event('blur', { bubbles: true }));
        outTimeInput.value = formatTime(endDate);
        outTimeInput.dispatchEvent(new Event('change', { bubbles: true }));
        outTimeInput.dispatchEvent(new Event('blur', { bubbles: true }));

        const commentTextarea = newWorkRow.querySelector(SELECTORS.commentTextarea);
        if (commentTextarea) {
            commentTextarea.value = event.subject;
            commentTextarea.dispatchEvent(new Event('input', { bubbles: true }));
            commentTextarea.dispatchEvent(new Event('change', { bubbles: true }));
        }

        if (event.project && event.project.trim() !== "") {
            await clickAndFillProject(newWorkRow, event.project);
        }
    }

    function retrieveTimesheetDates() {
        if (typeof window.model === 'object' && window.model.details && window.model.details.dateRange) {
            const dateObjToApiFormat = (dateObj) => `${dateObj.year}-${String(dateObj.month).padStart(2, '0')}-${String(dateObj.day).padStart(2, '0')}`;
            return {
                startDate: dateObjToApiFormat(window.model.details.dateRange.startDate),
                endDate: dateObjToApiFormat(window.model.details.dateRange.endDate)
            };
        }
        return null;
    }

    async function fetchAndFillData(button) {
        button.textContent = 'Fetching Dates...';
        button.disabled = true;
        let dates;
        try {
            dates = unsafeWindow[SCRIPT_ID + 'RetrieveDates']();
            if (!dates) throw new Error('Retrieved null dates. The window.model object is either missing or incomplete.');
        } catch (error) {
            handleError(error, button);
            return;
        }

        const apiUrl = `http://localhost:8000/calendar?from=${dates.startDate}&to=${dates.endDate}`;
        button.textContent = 'Fetching Calendar...';

        GM_xmlhttpRequest({
            method: 'GET',
            url: apiUrl,
            onload: async (response) => {
                try {
                    if (response.status >= 200 && response.status < 300) {
                        const calendarData = JSON.parse(response.responseText);
                        if (calendarData.length === 0) {
                            alert('No calendar events found for this timesheet period.');
                            button.textContent = 'Fetch Calendar Data';
                            button.disabled = false;
                            return;
                        }

                        for (const [index, event] of calendarData.entries()) {
                            button.textContent = `Populating ${index + 1}/${calendarData.length}...`;
                            if (event.subject === "Break Time") {
                                await addBreakToReplicon(event);
                            } else {
                                await addWorkEventToReplicon(event);
                            }
                            await new Promise(resolve => setTimeout(resolve, 1000));
                        }
                        alert('Successfully populated timesheet from calendar data!');
                        button.textContent = 'Done!';
                    } else {
                        throw new Error(`Server responded with status: ${response.status} ${response.statusText}`);
                    }
                } catch (err) {
                    handleError(err, button);
                }
            },
            onerror: (error) => {
                handleError(new Error(`Failed to fetch from localhost. Is your local server running?`), button);
            }
        });
    }

    function handleError(error, button) {
        console.error(`${SCRIPT_ID} Error:`, error.message);
        alert(`An error occurred: ${error.message}`);
        button.textContent = 'Error!';
        setTimeout(() => {
            button.textContent = 'Fetch Calendar Data';
            button.disabled = false;
        }, 5000);
    }

    function injectDateRetrievalScript() {
        const script = document.createElement('script');
        script.textContent = `(function() { window['${SCRIPT_ID}RetrieveDates'] = ${retrieveTimesheetDates.toString()}; })();`;
        document.head.appendChild(script);
        document.head.removeChild(script);
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
    injectDateRetrievalScript();

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

    const observer = new MutationObserver((mutationsList, obs) => {
        const container = document.querySelector(SELECTORS.buttonContainer);
        if (container && !document.getElementById(BUTTON_ID)) {
            injectButton(container);
            obs.disconnect();
        }
    });

    observer.observe(document.body, { childList: true, subtree: true });

})();