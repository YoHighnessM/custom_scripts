# Replacing Dashboard Table Scraping with Direct API Extraction

## Purpose

This guide explains how to update the existing Google Sheets + Google
Apps Script + Tampermonkey workflow so that Tampermonkey retrieves case
timestamps from the dashboard's JSON endpoint instead of reading the
rendered HTML table.

The discovered endpoint is:

``` http
GET https://api.dashboard.tech24et.com/api/callentries
```

A search for case `19311` produced a request equivalent to:

``` http
GET https://api.dashboard.tech24et.com/api/callentries?limit=20&search=19311&callstatus=&approvalstatus=&start_date=&end_date=&active=&bank=&branch=&district=
```

The response is JSON with a `data` array. Each record can contain:

-   `callentry_id`
-   `assigned_at`
-   `completed_at`
-   `created_at`
-   `updated_at`
-   `repeated_callentries_dates`

**Important:** This is an internal company endpoint. Only use it if your
company authorizes this automation and the account you use is permitted
to access these records. Do not try to bypass authentication, role
restrictions, or other access controls. Do not hardcode or share session
cookies, bearer tokens, or CSRF tokens.

------------------------------------------------------------------------

## 1. Understand the new workflow

Keep the existing Google Apps Script and Tampermonkey communication.
Replace only the part that searches the dashboard and scrapes the table.

``` text
Google Sheets
    |
    v
Google Apps Script
    |  Sends pending case IDs to the browser-side userscript
    v
Tampermonkey running on the logged-in dashboard
    |
    |  GET /api/callentries?search=<case ID>
    v
Company dashboard API
    |
    |  JSON response
    v
Tampermonkey validates and extracts timestamp fields
    |
    v
Google Apps Script receives structured results
    |
    v
Google Sheets is updated
```

The browser remains logged in. The userscript makes a request to the API
using the authenticated browser session. The Google Apps Script should
not receive or store the dashboard's cookies or tokens.

## 2. First test the request in the browser

Before editing either script, verify that the request works
independently.

1.  Log in to the company dashboard normally.
2.  Open Developer Tools (`F12`).
3.  Select **Console**.
4.  Run the test below.
5.  Check whether the returned record has `callentry_id: 19311` and the
    expected timestamps.

``` javascript
async function testCallentryApi(caseId) {
  const params = new URLSearchParams({
    limit: "20",
    search: String(caseId),
    callstatus: "",
    approvalstatus: "",
    start_date: "",
    end_date: "",
    active: "",
    bank: "",
    branch: "",
    district: ""
  });

  const url =
    `https://api.dashboard.tech24et.com/api/callentries?${params}`;

  const response = await fetch(url, {
    method: "GET",
    credentials: "include",
    headers: {
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    throw new Error(`HTTP ${response.status}: ${response.statusText}`);
  }

  return await response.json();
}

testCallentryApi(19311)
  .then(result => console.log("API result:", result))
  .catch(error => console.error("API test failed:", error));
```

### Expected result

The returned object should resemble:

``` javascript
{
  data: [
    {
      callentry_id: 19311,
      assigned_at: "2026-10-29T11:56:04.000000Z",
      completed_at: "2026-10-29 05:29 PM",
      created_at: "2026-10-29 02:55 PM",
      updated_at: "2026-10-29 06:10 PM",
      repeated_callentries_dates: [
        "2026-10-08 03:02 PM",
        "2026-10-29 02:55 PM"
      ]
    }
  ],
  meta: {
    total: 1,
    current_page: 1,
    last_page: 1
  }
}
```

This is an illustrative subset, not the complete response.

### If the test fails

-   **401 or 403:** The browser session may have expired, or the account
    may not have permission. Log in normally and check access with your
    company administrator. Do not copy cookies into the script.
-   **CORS error:** The browser is blocking the cross-origin request.
    Check the exact request's CORS headers and use the approved
    userscript request mechanism described in section 4. Do not disable
    browser security.
-   **404:** Verify the endpoint and path against the current Network
    request.
-   **200 but empty `data`:** Confirm the case ID and whether the
    dashboard applies additional filters.
-   **429:** The server is rate-limiting requests. Slow down and ask the
    dashboard administrator for an approved request rate.

A successful request in DevTools Console is a useful first check, but it
does not guarantee that a userscript's isolated execution context will
behave identically.

------------------------------------------------------------------------

## 3. Understand the query parameters

The observed request used these parameters:

  Parameter          Observed value          Purpose
  ------------------ ----------------------- ------------------------------------
  `limit`            `20`                    Maximum records requested per page
  `search`           Case ID, e.g. `19311`   Search term
  `callstatus`       empty                   Status filter
  `approvalstatus`   empty                   Approval filter
  `start_date`       empty                   Start-date filter
  `end_date`         empty                   End-date filter
  `active`           empty                   Active filter
  `bank`             empty                   Bank filter
  `branch`           empty                   Branch filter
  `district`         empty                   District filter

These are the parameters visible in the captured request. Their exact
filtering semantics should be confirmed with the dashboard owner or by
comparing authorized UI searches. Keep the same values the dashboard
sends unless you have confirmed that changing them is safe.

The endpoint appears to use pagination. The example response had
`meta.total: 1` and `meta.last_page: 1`. Do not assume every search will
return exactly one record.

------------------------------------------------------------------------

## 4. Add API extraction to Tampermonkey

### 4.1 Update the userscript metadata

At the top of the Tampermonkey script, check the metadata block.

For example:

``` javascript
// ==UserScript==
// @name         Company Dashboard Case Extractor
// @namespace    tech24-case-extractor
// @version      2.0.0
// @match        https://dashboard.tech24et.com/*
// @grant        GM_xmlhttpRequest
// @connect      api.dashboard.tech24et.com
// ==/UserScript==
```

Keep your existing `@match` and other metadata if they are already
correct. Add only the required grant and host permission.

`GM_xmlhttpRequest` is an alternative if normal `fetch()` is blocked by
CORS in the userscript context. It is a privileged userscript request,
so it should only be used for this explicitly authorized host. Do not
use `@connect *`.

### 4.2 Preferred first attempt: normal fetch

If the Console test succeeded, try this function from the userscript:

``` javascript
async function fetchCaseFromApi(caseId) {
  const params = new URLSearchParams({
    limit: "20",
    search: String(caseId),
    callstatus: "",
    approvalstatus: "",
    start_date: "",
    end_date: "",
    active: "",
    bank: "",
    branch: "",
    district: ""
  });

  const url =
    `https://api.dashboard.tech24et.com/api/callentries?${params}`;

  const response = await fetch(url, {
    method: "GET",
    credentials: "include",
    headers: {
      Accept: "application/json"
    }
  });

  if (!response.ok) {
    throw new Error(`API returned HTTP ${response.status}`);
  }

  const result = await response.json();
  return extractCaseFields(result, caseId);
}
```

If this works in Tampermonkey, you do not need `GM_xmlhttpRequest`.

### 4.3 Fallback: Tampermonkey GM_xmlhttpRequest

If `fetch()` fails specifically because of userscript CORS restrictions,
use this version. It relies on the browser's existing authenticated
session; do not set a manually copied `Cookie` header.

``` javascript
function gmGetJson(url) {
  return new Promise((resolve, reject) => {
    GM_xmlhttpRequest({
      method: "GET",
      url,
      headers: {
        Accept: "application/json"
      },
      onload: response => {
        if (response.status < 200 || response.status >= 300) {
          reject(new Error(`API returned HTTP ${response.status}`));
          return;
        }

        try {
          resolve(JSON.parse(response.responseText));
        } catch (error) {
          reject(new Error("API response was not valid JSON"));
        }
      },
      onerror: () => reject(new Error("Network request failed")),
      ontimeout: () => reject(new Error("API request timed out"))
    });
  });
}

async function fetchCaseFromApiGM(caseId) {
  const params = new URLSearchParams({
    limit: "20",
    search: String(caseId),
    callstatus: "",
    approvalstatus: "",
    start_date: "",
    end_date: "",
    active: "",
    bank: "",
    branch: "",
    district: ""
  });

  const url =
    `https://api.dashboard.tech24et.com/api/callentries?${params}`;

  const result = await gmGetJson(url);
  return extractCaseFields(result, caseId);
}
```

**Authentication note:** Whether `GM_xmlhttpRequest` sends the
dashboard's cookies depends on the userscript manager, its
configuration, browser behavior, and cookie scope. Test it while logged
in. If it does not authenticate, do not copy session cookies into code.
Prefer the working credentialed `fetch()` from the dashboard context or
ask the dashboard administrator for an approved integration method.

------------------------------------------------------------------------

## 5. Validate and extract the response

Use a separate function to validate the API response and select only the
required fields.

``` javascript
function extractCaseFields(result, requestedCaseId) {
  if (!result || !Array.isArray(result.data)) {
    throw new Error("Unexpected API response: data array missing");
  }

  const requestedId = String(requestedCaseId);

  const record = result.data.find(
    item => String(item.callentry_id) === requestedId
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
    caseId: String(record.callentry_id),
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
}
```

Why validate the ID? A search endpoint may return multiple records or
partial matches. Never assume the first item is necessarily the exact
case requested.

### Timestamp handling

The sample response contains different formats:

-   `assigned_at`: ISO UTC string, e.g. `2026-10-29T11:56:04.000000Z`
-   `completed_at`, `created_at`, `updated_at`: formatted strings,
    e.g. `2026-10-29 05:29 PM`
-   `repeated_callentries_dates`: array of formatted date/time strings

Initially, pass these strings through unchanged. Do not parse and
reformat them until you have confirmed the intended spreadsheet timezone
and display format. In particular, converting a UTC timestamp to local
time can change the displayed hour.

------------------------------------------------------------------------

## 6. Process multiple case IDs safely

Do not fire hundreds of requests simultaneously. Start with sequential
processing and a small delay. This is easier to debug and less likely to
overload the internal service.

``` javascript
const sleep = ms =>
  new Promise(resolve => setTimeout(resolve, ms));

async function fetchManyCases(caseIds) {
  const results = [];

  for (const caseId of caseIds) {
    try {
      const result = await fetchCaseFromApi(caseId);
      results.push(result);
    } catch (error) {
      results.push({
        caseId: String(caseId),
        found: false,
        error: error.message
      });
    }

    // Start conservatively. Confirm an acceptable rate with the
    // dashboard administrator before increasing throughput.
    await sleep(1000);
  }

  return results;
}
```

If using the GM request fallback, replace `fetchCaseFromApi` with
`fetchCaseFromApiGM`.

Consider these cases in your final implementation:

-   Empty or invalid case IDs.
-   Case not found.
-   More than one returned record.
-   API errors and expired sessions.
-   Request timeouts.
-   Rate limiting (`429`).
-   Partial batch completion.
-   Duplicate IDs in the spreadsheet.

For `429` responses, pause and retry with an increasing delay rather
than immediately repeating the request.

### Pagination

The sample response included:

``` json
{
  "meta": {
    "current_page": 1,
    "last_page": 1,
    "per_page": 20,
    "total": 1
  }
}
```

For a case-ID search, one exact result is expected in the example, but
the implementation should still validate the exact ID. If the endpoint
returns multiple pages or the exact match is not in the first page,
follow the API's documented pagination behavior. Do not invent page
parameters: confirm them from the dashboard's Network requests.

------------------------------------------------------------------------

## 7. Connect this to your existing Google Apps Script

You already have a working Apps Script ↔ Tampermonkey communication
mechanism. Keep it.

The clean division of responsibilities should be:

### Google Apps Script

1.  Read pending case IDs from the spreadsheet.
2.  Send a batch/job to the browser-side userscript using your existing
    mechanism.
3.  Receive the extracted results.
4.  Match each result to its spreadsheet row using the case ID.
5.  Write only the required timestamp columns.
6.  Mark each case as processed, not found, or failed.

### Tampermonkey

1.  Receive the job from Apps Script.
2.  Validate the case ID.
3.  Call the internal API.
4.  Validate that the returned record's `callentry_id` matches the
    requested ID.
5.  Extract only the required timestamp fields.
6.  Send structured results back through your existing communication
    mechanism.

### Suggested result object

Use a predictable object shape between Tampermonkey and Apps Script:

``` javascript
{
  caseId: "19311",
  found: true,
  assignedAt: "2026-10-29T11:56:04.000000Z",
  completedAt: "2026-10-29 05:29 PM",
  createdAt: "2026-10-29 02:55 PM",
  updatedAt: "2026-10-29 06:10 PM",
  repeatedDates: [
    "2026-10-08 03:02 PM",
    "2026-10-29 02:55 PM"
  ],
  error: ""
}
```

For a missing case:

``` javascript
{
  caseId: "19311",
  found: false,
  assignedAt: "",
  completedAt: "",
  createdAt: "",
  updatedAt: "",
  repeatedDates: [],
  error: "Case not found"
}
```

The Apps Script should update rows by case ID, not by assuming that
results arrive in the same order as the input list.

------------------------------------------------------------------------

## 8. Security and reliability checklist

-   [ ] Use only an authorized company account and approved automation.
-   [ ] Keep dashboard authentication in the browser.
-   [ ] Never hardcode cookies, session IDs, CSRF tokens, or bearer
    tokens.
-   [ ] Never send authentication values to Google Sheets.
-   [ ] Do not log full request headers or complete records containing
    unrelated personal data.
-   [ ] Restrict Tampermonkey's `@connect` permission to the exact API
    hostname.
-   [ ] Extract only the fields needed by the spreadsheet.
-   [ ] Validate every returned case ID.
-   [ ] Handle empty responses and HTTP errors.
-   [ ] Use conservative request pacing.
-   [ ] Stop processing if authentication expires or access is denied.
-   [ ] Avoid storing sensitive response data in browser local storage
    or script logs.
-   [ ] Confirm with the dashboard owner that automated use of this
    endpoint is permitted.

The API may be an internal implementation detail and can change without
notice. If this workflow becomes business-critical, request a
documented, supported API or service account from the dashboard owner.

------------------------------------------------------------------------

## 9. Recommended implementation order

Do the migration in small steps. Keep a backup of the current working
scripts.

### Stage 1 --- Confirm API access

-   [ ] Run the Console test while logged in.
-   [ ] Confirm case `19311` returns the expected record.
-   [ ] Confirm the timestamp values match the dashboard.
-   [ ] Test one case that has no result, if available.

### Stage 2 --- Add extraction functions

-   [ ] Add `fetchCaseFromApi`.
-   [ ] Add `extractCaseFields`.
-   [ ] Test one ID from the Tampermonkey console.
-   [ ] Test several IDs sequentially.
-   [ ] Verify error handling and rate limiting.

### Stage 3 --- Connect to the existing Apps Script

-   [ ] Keep the existing job/request mechanism.
-   [ ] Replace the DOM scraping function with the API extraction
    function.
-   [ ] Keep the existing result-return mechanism.
-   [ ] Match results to sheet rows by case ID.

### Stage 4 --- Validate before rollout

-   [ ] Compare results for at least 10 cases against the dashboard.
-   [ ] Verify timezone behavior.
-   [ ] Verify missing and duplicate cases.
-   [ ] Verify that a session expiry stops the job safely.
-   [ ] Run a small batch before processing the full sheet.

### Stage 5 --- Retire the old scraper

Only remove the old search-bar and table-scraping code after the new
implementation has been tested successfully and the results have been
reconciled.

------------------------------------------------------------------------

## 10. Information to confirm before finalizing the scripts

These details are not established by the captured response and should be
confirmed:

1.  Does the API search return exact matches only, or can it return
    partial matches?
2.  Does the API require any additional request headers beyond
    `Accept: application/json`?
3.  Does normal `fetch()` work from the dashboard page's DevTools
    Console?
4.  Does normal `fetch()` also work inside the Tampermonkey userscript?
5.  Which timestamp columns does the spreadsheet actually need?
6.  How does the current Apps Script deliver jobs to Tampermonkey and
    receive results?
7.  What request rate is acceptable to the dashboard owner?

**Recommended approach:** preserve the current Google Apps Script ↔
Tampermonkey bridge, use the authenticated browser session to call the
discovered JSON endpoint, validate the returned case ID, and send only
the required timestamp fields back to Sheets.
