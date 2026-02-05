/**
 * Code.gs - Registration System Backend
 * -------------------------------------
 * 1. Copy ALL content below.
 * 2. Paste into your Google Apps Script editor (replace everything).
 * 3. Update 'VERIFY_BASE_URL' and 'ADMIN_KEY' constants.
 * 4. Save.
 * 5. Add an Installable Trigger for 'installedOnEdit':
 *    - Go to Triggers -> Add Trigger -> installedOnEdit -> From spreadsheet -> On edit.
 * 6. Deploy -> New Version -> Web App -> Execute as Me -> Access: Anyone.
 */

// --- CONSTANTS ---
var PAYMENT_PROOFS_FOLDER_NAME = 'Payment-Proofs';
var TICKET_FOLDER_NAME = 'Event-Tickets';
var SHEET_NAME = 'Registrations';

// UPDATE THEM
var VERIFY_BASE_URL = 'https://bis-registration.netlify.app/verify.html';
var ADMIN_KEY = 'BIScusat';

// --- HTTP HANDLERS ---

function doPost(e) {
    var lock = LockService.getScriptLock();
    lock.tryLock(10000);

    try {
        if (!e || !e.postData) return jsonResponse({ status: "error", message: "Missing request body" });

        var data;
        try {
            data = JSON.parse(e.postData.contents);
        } catch (err) {
            return jsonResponse({ status: "error", message: "Invalid JSON" });
        }

        // STRICT ROUTING
        if (data.action === 'checkin') {
            return handleCheckIn(data);
        } else if (data.action === 'manual_checkin') {
            return handleManualCheckIn(data);
        } else if (data.action === 'delete') {
            return handleDelete(data);
        } else if (data.action === 'update_status') {
            return handleUpdateStatus(data);
        } else if (data.action === 'markAttendance') {
            return handleMarkAttendance(data);
        } else if (data.action === 'register' || !data.action) {
            // Default to registration ONLY if action is 'register' or missing (legacy form)
            return handleRegistration(data);
        } else {
            return jsonResponse({ result: "error", message: "Invalid Action: " + data.action });
        }

    } catch (error) {
        return jsonResponse({ status: "error", message: "Server Error: " + error.toString() });
    } finally {
        lock.releaseLock();
    }
}

function doGet(e) {
    var p = e.parameter;
    var action = (p.action || "").toLowerCase(); // Case-insensitive

    try {
        if (action === 'get') {
            return handleGetTicket(p.ticket, p.t);
        }
        if (action === 'list_all') {
            return handleListAll();
        }
        if (action === 'stats') {
            return handleStats(p);
        }
        if (action === 'proxyimage') {
            return handleProxyImage(p);
        }
        return jsonResponse({ status: "error", message: "Unknown action" });
    } catch (error) {
        return jsonResponse({ status: "error", message: error.toString() });
    }
}

function doOptions(e) {
    var output = ContentService.createTextOutput("");
    output.setMimeType(ContentService.MimeType.JSON);
    return addCorsHeaders(output);
}

// --- CORE IMPLEMENTATION ---

function handleRegistration(data) {
    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    if (!headers.length) return jsonResponse({ status: "error", message: "Headers missing" });

    // Robust Header Lookup (Case-Insensitive)
    var getIdx = function (name) {
        var lower = name.toLowerCase().trim();
        for (var i = 0; i < headers.length; i++) {
            if (String(headers[i]).toLowerCase().trim() === lower) return i;
        }
        return -1;
    };

    var fileUrl = "";
    if (data.screenshot) {
        fileUrl = saveToDrive_(data.screenshot, data.mimeType, data.fileName || "proof.png", PAYMENT_PROOFS_FOLDER_NAME);
    }

    // Format Timestamp as String to ensure it stays static
    var now = new Date();
    var timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

    var rowData = new Array(headers.length).fill(""); // Initialize empty row

    // Helper to set value by header name safely
    var setVal = function (headerName, value) {
        var idx = getIdx(headerName);
        if (idx > -1) rowData[idx] = value;
    };

    setVal("Timestamp", timestampStr);
    setVal("Name", data.name);
    setVal("Email", data.email);
    setVal("Phone", data.phone);
    setVal("Transaction ID", data.transactionId || "");
    setVal("Payment Proof URL", fileUrl);
    setVal("Status", "Pending");
    setVal("Ticket ID", "");
    setVal("Ticket Token", "");
    setVal("Ticket Link", "");
    setVal("Ticket QR File URL", "");
    setVal("Ticket Sent", "No");
    setVal("Checked In", "No");

    sheet.appendRow(rowData);

    return jsonResponse({ status: "success", message: "Registration submitted successfully" });
}

function handleCheckIn(data) {
    if (data.adminKey !== ADMIN_KEY) return jsonResponse({ status: "error", message: "Invalid Admin Key" });

    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    var values = sheet.getDataRange().getValues();

    var idxTicket = headers.indexOf("Ticket ID");
    var idxToken = headers.indexOf("Ticket Token");
    var idxChecked = headers.indexOf("Checked In");

    if (idxTicket === -1 || idxToken === -1 || idxChecked === -1)
        return jsonResponse({ status: "error", message: "Missing columns" });

    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (row[idxTicket] === data.ticket && row[idxToken] === data.t) {
            var idxStatus = headers.indexOf("Status");
            if (idxStatus > -1 && String(row[idxStatus]).toLowerCase() === 'rejected') {
                return jsonResponse({ status: "error", message: "Creation Denied: Rejected Ticket" });
            }

            if (String(row[idxChecked]).toLowerCase() === 'yes') {
                return jsonResponse({ status: "error", message: "Already Checked In" });
            }
            sheet.getRange(i + 1, idxChecked + 1).setValue("Yes");

            // --- AUTO MARK DAY 1 ATTENDANCE ---
            var colDay1 = "Day 1 Attendance";
            var idxDay1 = headers.indexOf(colDay1);
            if (idxDay1 > -1) {
                var currentDay1 = row[idxDay1];
                if (!currentDay1 || String(currentDay1).length < 2) {
                    var now = new Date();
                    var timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                    sheet.getRange(i + 1, idxDay1 + 1).setValue(timestampStr);
                }
            }
            // ----------------------------------

            return jsonResponse({
                status: "success",
                message: "Check-in Successful for " + row[headers.indexOf("Name")]
            });
        }
    }
    return jsonResponse({ status: "error", message: "Ticket not found" });
}

function handleGetTicket(ticketId, token) {
    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    var values = sheet.getDataRange().getValues();

    var idxTicket = headers.indexOf("Ticket ID");
    var idxToken = headers.indexOf("Ticket Token");

    if (!ticketId || !token) return jsonResponse({ status: "error", message: "Missing params" });

    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (row[idxTicket] === ticketId && row[idxToken] === token) {
            var ticketObj = {};
            headers.forEach(function (h, k) { ticketObj[h] = row[k]; });
            return jsonResponse({ status: "success", ticket: ticketObj });
        }
    }
    return jsonResponse({ status: "error", message: "Ticket not found" });
}

function handleListAll() {
    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    var values = sheet.getDataRange().getValues();
    var rows = [];
    for (var i = 1; i < values.length; i++) {
        var rowObj = {};
        for (var h = 0; h < headers.length; h++) {
            rowObj[headers[h]] = values[i][h];
        }
        rows.push(rowObj);
    }
    return jsonResponse({ status: "success", rows: rows });
}

function handleDelete(data) {
    if (data.adminKey !== ADMIN_KEY) return jsonResponse({ status: "error", message: "Unauthorized" });

    var sheet = getSheet_();
    var row = parseInt(data.row);

    if (isNaN(row) || row < 2) {
        return jsonResponse({ status: "error", message: "Invalid Row Index" });
    }

    try {
        // Optional: specific check to ensure we aren't deleting wrong row?
        // e.g. check email matches data.email if provided.
        // For now, direct deletion by row ID which comes from fresh stats.
        sheet.deleteRow(row);
        return jsonResponse({ status: "success", message: "Deleted row " + row });
    } catch (e) {
        return jsonResponse({ status: "error", message: "Delete Exception: " + e.toString() });
    }
}

function handleUpdateStatus(data) {
    if (data.adminKey !== ADMIN_KEY) return jsonResponse({ status: "error", message: "Unauthorized" });

    var sheet = getSheet_();
    var row = parseInt(data.row);
    var newStatus = data.status;

    if (isNaN(row) || row < 2) return jsonResponse({ status: "error", message: "Invalid Row" });
    if (!newStatus) return jsonResponse({ status: "error", message: "Missing Status" });

    var headers = getHeaders_(sheet);
    var idxStatus = headers.indexOf("Status");
    if (idxStatus === -1) return jsonResponse({ status: "error", message: "Status column not found" });

    sheet.getRange(row, idxStatus + 1).setValue(newStatus);

    // TRIGGER TICKET GENERATION IF APPROVED
    if (newStatus === 'Approved') {
        try {
            // Optional: Check if already sent to avoid double send? 
            // For now, let's assume if Admin sets to Approved, they want a ticket.
            // But checking 'Ticket Sent' is safer to avoid spam if clicked multiple times.
            var idxSent = headers.indexOf("Ticket Sent");
            var isSent = (idxSent > -1) ? sheet.getRange(row, idxSent + 1).getValue() : "No";

            if (String(isSent).toLowerCase() !== 'yes') {
                processTicketGeneration_(sheet, row, headers);
                return jsonResponse({ status: "success", message: "Updated to Approved & Ticket Generated" });
            } else {
                return jsonResponse({ status: "success", message: "Updated to Approved (Ticket already sent)" });
            }
        } catch (e) {
            // If generation fails, still return success for the update but warn
            return jsonResponse({ status: "success", message: "Updated to Approved but Ticket Error: " + e.toString() });
        }
    }

    return jsonResponse({ status: "success", message: "Updated status to " + newStatus });
}

function handleManualCheckIn(data) {
    if (data.adminKey !== ADMIN_KEY) return jsonResponse({ status: "error", message: "Unauthorized" });

    var ticketId = data.ticketId;
    if (!ticketId) return jsonResponse({ status: "error", message: "Missing Ticket ID" });

    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    var values = sheet.getDataRange().getValues();

    var idxTicket = headers.indexOf("Ticket ID");
    var idxChecked = headers.indexOf("Checked In");
    var idxName = headers.indexOf("Name");

    if (idxTicket === -1 || idxChecked === -1)
        return jsonResponse({ status: "error", message: "Missing columns" });

    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (String(row[idxTicket]).trim() === ticketId.trim()) {
            var idxStatus = headers.indexOf("Status");
            if (idxStatus > -1 && String(row[idxStatus]).toLowerCase() === 'rejected') {
                return jsonResponse({ status: "error", message: "Check-in Denied: Rejected Status" });
            }

            if (String(row[idxChecked]).toLowerCase() === 'yes') {
                return jsonResponse({ status: "error", message: "Already Checked In" });
            }
            sheet.getRange(i + 1, idxChecked + 1).setValue("Yes");

            // --- AUTO MARK DAY 1 ATTENDANCE (Added for Manual Check-in) ---
            var colDay1 = "Day 1 Attendance";
            var idxDay1 = headers.indexOf(colDay1);
            if (idxDay1 > -1) {
                var currentDay1 = row[idxDay1];
                if (!currentDay1 || String(currentDay1).length < 2) {
                    var now = new Date();
                    var timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
                    sheet.getRange(i + 1, idxDay1 + 1).setValue(timestampStr);
                }
            }
            // -------------------------------------------------------------
            return jsonResponse({
                status: "success",
                message: "Checked In: " + row[idxName]
            });
        }
    }
    return jsonResponse({ status: "error", message: "Ticket ID Not Found" });
}

function handleMarkAttendance(data) {
    if (data.adminKey !== ADMIN_KEY) return jsonResponse({ status: "error", message: "Unauthorized: Invalid key" });

    var ticketId = data.ticket;
    var token = data.token;
    var day = data.day; // "day1", "day2", "day3"

    if (!ticketId || !token || !day) return jsonResponse({ status: "error", message: "Missing info" });

    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    var values = sheet.getDataRange().getValues();

    var idxTicket = headers.indexOf("Ticket ID");
    var idxToken = headers.indexOf("Ticket Token");
    var idxStatus = headers.indexOf("Status");

    // Map day to column name
    var colName = "";
    if (day === "day1") colName = "Day 1 Attendance";
    else if (day === "day2") colName = "Day 2 Attendance";
    else if (day === "day3") colName = "Day 3 Attendance";
    else return jsonResponse({ status: "error", message: "Invalid Day" });

    var idxDay = headers.indexOf(colName);

    if (idxTicket === -1 || idxToken === -1 || idxDay === -1)
        return jsonResponse({ status: "error", message: "Missing columns (Check sheet headers)" });

    for (var i = 1; i < values.length; i++) {
        var row = values[i];
        if (String(row[idxTicket]).trim() === String(ticketId).trim()) {
            // Verify Token
            if (String(row[idxToken]).trim() !== String(token).trim()) {
                return jsonResponse({ status: "error", message: "Invalid Token" });
            }
            // Verify Approved
            if (String(row[idxStatus]).toLowerCase() !== "approved") {
                return jsonResponse({ status: "error", message: "Ticket not Approved" });
            }

            // Check if already marked
            var currentVal = row[idxDay];
            if (currentVal && String(currentVal).length > 2) {
                return jsonResponse({
                    success: false,
                    message: "Attendance already marked for " + colName
                });
            }

            // Mark it
            var now = new Date();
            var timestampStr = Utilities.formatDate(now, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
            sheet.getRange(i + 1, idxDay + 1).setValue(timestampStr);

            return jsonResponse({
                success: true,
                message: colName + " marked successfully",
                timestamp: timestampStr,
                name: row[headers.indexOf("Name")]
            });
        }
    }

    return jsonResponse({ status: "error", message: "Ticket not found" });
}

// --- ANALYTICS HANDLER ---

function handleStats(params) {
    if (params.adminKey !== ADMIN_KEY) {
        return jsonResponse({ success: false, message: "Unauthorized: Invalid Admin Key" });
    }

    var sheet = getSheet_();
    if (sheet.getLastColumn() === 0) {
        return jsonResponse({ success: true, meta: { total: 0, approved: 0, pending: 0, checkedIn: 0 }, timeseries: [], recent: [] });
    }
    var headers = getHeaders_(sheet);
    var data = sheet.getDataRange().getValues(); // Header is row 0
    var rows = data.slice(1);

    var lastDays = parseInt(params.lastDays) || 30;
    var now = new Date();
    var cutoffDate = new Date();
    cutoffDate.setDate(now.getDate() - lastDays);

    // Indices - Robust Lookup
    var getIdx = function (name) {
        var lower = name.toLowerCase().trim();
        for (var i = 0; i < headers.length; i++) {
            if (String(headers[i]).toLowerCase().trim() === lower) return i;
        }
        return -1;
    };

    var iTimestamp = getIdx("Timestamp");
    // If not found, try generic 'timestamp' or 'date' or just 0
    if (iTimestamp === -1) iTimestamp = getIdx("timestamp");

    var iStatus = getIdx("Status");
    var iCheckedIn = getIdx("Checked In");
    var iName = getIdx("Name");
    var iEmail = getIdx("Email");
    var iPhone = getIdx("Phone");
    var iTxn = getIdx("Transaction ID");
    var iProof = getIdx("Payment Proof URL");
    var iTicketId = getIdx("Ticket ID");

    var iDay1 = getIdx("Day 1 Attendance");
    var iDay2 = getIdx("Day 2 Attendance");
    var iDay3 = getIdx("Day 3 Attendance");

    // Aggregates
    var total = 0, approved = 0, pending = 0, checkedIn = 0;
    var d1Count = 0, d2Count = 0, d3Count = 0;
    var timeseriesMap = {};
    var recent = [];

    // Initialize timeseries map
    for (var d = 0; d < lastDays; d++) {
        var day = new Date();
        day.setDate(now.getDate() - d);
        var key = formatDateISO_(day);
        timeseriesMap[key] = { date: key, registrations: 0, approved: 0, checkedIn: 0 };
    }

    // Process rows
    for (var i = 0; i < rows.length; i++) {
        var row = rows[i];

        // Robust Date Parsing
        var tsVal = (iTimestamp > -1) ? row[iTimestamp] : null;
        var ts = null;

        if (tsVal) {
            if (tsVal instanceof Date) {
                ts = tsVal;
            } else {
                ts = new Date(tsVal);
            }
        }

        // If INVALID date, keep as null to indicate "Unknown"
        if (ts && isNaN(ts.getTime())) {
            ts = null;
        }

        var status = (iStatus > -1 && row[iStatus] !== undefined) ? String(row[iStatus]) : "";
        var isApproved = status.toLowerCase() === 'approved';
        var isCheckedVal = (iCheckedIn > -1 && row[iCheckedIn] !== undefined) ? String(row[iCheckedIn]) : "No";
        var isChecked = isCheckedVal.toLowerCase() === 'yes';

        var d1Val = (iDay1 > -1) ? row[iDay1] : "";
        var d2Val = (iDay2 > -1) ? row[iDay2] : "";
        var d3Val = (iDay3 > -1) ? row[iDay3] : "";

        var hasD1 = d1Val && String(d1Val).length > 2;
        var hasD2 = d2Val && String(d2Val).length > 2;
        var hasD3 = d3Val && String(d3Val).length > 2;

        // Global Totals
        total++;
        if (isApproved) approved++;
        if (status.toLowerCase() === 'pending') pending++;
        if (isChecked) checkedIn++;

        if (hasD1) d1Count++;
        if (hasD2) d2Count++;
        if (hasD3) d3Count++;

        // Timeseries & Recent
        // Only add to timeseries if we have a valid timestamp
        if (ts) {
            if (ts >= cutoffDate) {
                var key = formatDateISO_(ts);
                if (!timeseriesMap[key]) timeseriesMap[key] = { date: key, registrations: 0, approved: 0, checkedIn: 0 };

                timeseriesMap[key].registrations++;
                if (isApproved) timeseriesMap[key].approved++;
                if (isChecked) timeseriesMap[key].checkedIn++;
            }
        }

        recent.push({
            Row: i + 2,
            Timestamp: ts, // Can be null
            Name: iName > -1 ? row[iName] : "N/A",
            Email: iEmail > -1 ? row[iEmail] : "",
            Phone: iPhone > -1 ? row[iPhone] : "",
            "Transaction ID": iTxn > -1 ? row[iTxn] : "",
            ProofURL: iProof > -1 ? row[iProof] : "",
            Status: status,
            TicketID: iTicketId > -1 ? row[iTicketId] : "",
            TicketToken: getIdx("Ticket Token") > -1 ? row[getIdx("Ticket Token")] : "",
            TicketLink: getIdx("Ticket Link") > -1 ? row[getIdx("Ticket Link")] : "",
            TicketQR: getIdx("Ticket QR File URL") > -1 ? row[getIdx("Ticket QR File URL")] : "",
            CheckedIn: isChecked ? "Yes" : "No",
            Day1: hasD1 ? d1Val : "",
            Day2: hasD2 ? d2Val : "",
            Day3: hasD3 ? d3Val : ""
        });
    }

    // Sort recent by date desc
    recent.sort(function (a, b) { return b.Timestamp - a.Timestamp; });

    // Convert maps to sorted arrays
    var timeseries = Object.keys(timeseriesMap).sort().map(function (k) { return timeseriesMap[k]; });

    // Recent Checked In List
    var recentCheckedIn = recent.filter(function (r) { return r.CheckedIn === "Yes"; }).slice(0, 20);

    return jsonResponse({
        success: true,
        debug: {
            sheetName: sheet.getName(),
            totalRows: rows.length,
            headersFound: headers,
            indices: { iTimestamp: iTimestamp, iStatus: iStatus, iCheckedIn: iCheckedIn },
            sampleRow: rows.length > 0 ? rows[0] : "Empty",
            rawTimestamp: rows.length > 0 ? rows[0][iTimestamp] : "N/A",
            sampleDateParsed: rows.length > 0 ? new Date(rows[0][iTimestamp]).toString() : "N/A"
        },
        meta: {
            total: total,
            approved: approved,
            pending: pending,
            checkedIn: checkedIn,
            d1: d1Count,
            d2: d2Count,
            d3: d3Count,
            conversionPct: total > 0 ? ((approved / total) * 100).toFixed(1) : 0,
            checkedInPct: approved > 0 ? ((checkedIn / approved) * 100).toFixed(1) : 0
        },
        timeseries: timeseries,
        breakdown: {
            paid: approved,
            pending: pending,
            unpaid: total - approved - pending
        },
        recent: recent.slice(0, 50), // Limit to 50 for performance
        recentCheckedIn: recentCheckedIn
    });
}

function formatDateISO_(date) {
    return Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

// --- TRIGGER LOGIC (INSTALLABLE) ---

function installedOnEdit(e) {
    var sheet = e.source.getActiveSheet();
    if (sheet.getName() !== SHEET_NAME) return;

    // Only handle single cell edits
    if (e.range.getNumRows() > 1 || e.range.getNumColumns() > 1) return;

    var headers = getHeaders_(sheet);
    var idxStatus = headers.indexOf("Status");

    // Check if Status column edited
    if (e.range.getColumn() === (idxStatus + 1) && e.range.getRow() > 1) {
        if (e.range.getValue() === 'Approved') {
            // Process
            processTicketGeneration_(sheet, e.range.getRow(), headers);
        }
    }
}

// --- MANUAL DEBUG FUNCTION ---

/**
 * Manually generate ticket for a specific row number.
 * Usage: Run this function from the GAS Editor -> Select 'generateTicketForRowManual' -> Run.
 */
function generateTicketForRowManual(rowNum) {
    Logger.log("--- Starting Manual Ticket Generation for Row: " + rowNum + " ---");

    // Default to row 2 if not provided (for testing)
    if (typeof rowNum === 'object') rowNum = 2; // Handle if called from UI trigger accidentally
    if (!rowNum) rowNum = 2;

    var sheet = getSheet_();
    var headers = getHeaders_(sheet);
    var values = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getValues()[0];

    var idxStatus = headers.indexOf("Status");
    var idxTicketSent = headers.indexOf("Ticket Sent");

    var status = values[idxStatus];
    var ticketSent = values[idxTicketSent];

    Logger.log("Current Status: " + status);
    Logger.log("Ticket Sent: " + ticketSent);

    if (status !== 'Approved') {
        Logger.log("Error: Status is not 'Approved'. Skipping.");
        return;
    }

    if (ticketSent === 'Yes') {
        Logger.log("Warning: Ticket already sent ('Yes'). Continuing anyway for manual test...");
        // return; // Uncomment to strict check
    }

    processTicketGeneration_(sheet, rowNum, headers);
    Logger.log("--- Manual Generation Complete ---");
}

function processTicketGeneration_(sheet, rowIndex, headers) {
    // Read fresh data
    var rowData = sheet.getRange(rowIndex, 1, 1, sheet.getLastColumn()).getValues()[0];

    var ticketId = generateTicketId();
    var ticketToken = generateToken();
    var ticketLink = VERIFY_BASE_URL + '?ticket=' + ticketId + '&t=' + ticketToken;

    Logger.log("Generated ID: " + ticketId);

    // QR Generation
    var qrBlob = generateQrBlob(ticketLink);
    var qrFileName = ticketId + ".png";
    var qrUrl = saveToDriveBlob_(qrBlob, TICKET_FOLDER_NAME, qrFileName);

    Logger.log("QR Saved. URL: " + qrUrl);

    // Update Sheet
    var setVal = function (header, val) {
        var col = headers.indexOf(header);
        if (col > -1) sheet.getRange(rowIndex, col + 1).setValue(val);
    };

    setVal("Ticket ID", ticketId);
    setVal("Ticket Token", ticketToken);
    setVal("Ticket Link", ticketLink);
    setVal("Ticket QR File URL", qrUrl);
    setVal("Ticket Sent", "Yes");

    Logger.log("Sheet Updated.");

    // Email
    var email = rowData[headers.indexOf("Email")];
    var name = rowData[headers.indexOf("Name")];
    sendTicketEmail(name, email, ticketId, ticketLink);
}

// --- QR GENERATION ---

function generateQrBlob(text) {
    try {
        // API: QrServer
        // endpoint: https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=...
        var apiUrl = "https://api.qrserver.com/v1/create-qr-code/?size=400x400&data=" + encodeURIComponent(text);

        var options = {
            "method": "get",
            "muteHttpExceptions": true,
            "validateHttpsCertificates": true,
            "timeoutSeconds": 30 // Explicit timeout
        };

        var response = UrlFetchApp.fetch(apiUrl, options);
        var rc = response.getResponseCode();

        if (rc !== 200) {
            throw new Error("QR API Failed. Code: " + rc + " Body: " + response.getContentText().substring(0, 100));
        }

        var blob = response.getBlob();
        blob.setName("ticket_qr.png");
        return blob;

    } catch (e) {
        Logger.log("Error in generateQrBlob: " + e.toString());
        throw e; // Rethrow to stop process
    }
}

// Fallback logic (Unused but kept for reference)
function generateQrBlob_googleChart_fallback(text) {
    try {
        var url = "https://chart.googleapis.com/chart?chs=400x400&cht=qr&chl=" + encodeURIComponent(text);
        var response = UrlFetchApp.fetch(url);
        if (response.getResponseCode() !== 200) throw new Error("Google Chart API failed");
        return response.getBlob().setName("qr.png");
    } catch (e) {
        Logger.log("Google Chart Fallback Error: " + e.toString());
        throw e;
    }
}


// --- HELPER FUNCTIONS ---

function getSheet_() {
    return SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME) || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
}

/**
 * Sends a professional ticket email to the attendee.
 */
function sendTicketEmail(name, email, ticketId, ticketLink) {
    if (!email) {
        Logger.log("sendTicketEmail: No email provided.");
        return;
    }

    var subject = "Your BIS AutoCAD Internship Ticket";

    // Professional HTML Email Template
    var htmlBody =
        '<div style="font-family: Arial, sans-serif; max-width: 600px; margin: 0 auto; padding: 20px; border: 1px solid #e0e0e0; border-radius: 8px; background-color: #ffffff;">' +
        '<h2 style="color: #333333; margin-top: 0;">Registration Approved!</h2>' +
        '<p style="color: #555555; font-size: 16px; line-height: 1.5;">Hello <strong>' + escapeHtml(name) + '</strong>,</p>' +
        '<p style="color: #555555; font-size: 16px; line-height: 1.5;">We are excited to confirm your registration for the BIS AutoCAD Internship. Your ticket has been generated successfully.</p>' +

        '<div style="background-color: #f8f9fa; border-top: 4px solid #4285f4; padding: 20px; margin: 20px 0; border-radius: 4px; text-align: center;">' +
        '<p style="margin: 0; color: #777777; font-size: 14px; text-transform: uppercase; letter-spacing: 1px;">Ticket ID</p>' +
        '<p style="margin: 5px 0 0 0; color: #333333; font-size: 24px; font-weight: bold; letter-spacing: 1px;">' + ticketId + '</p>' +
        '</div>' +

        '<div style="text-align: center; margin: 30px 0;">' +
        '<a href="' + ticketLink + '" style="background-color: #4285f4; color: #ffffff; padding: 15px 30px; text-decoration: none; border-radius: 5px; font-weight: bold; font-size: 18px; display: inline-block;">View Your Ticket</a>' +
        '</div>' +

        '<p style="color: #555555; font-size: 14px; text-align: center;">' +
        'or copy this link:<br>' +
        '<a href="' + ticketLink + '" style="color: #4285f4; word-break: break-all;">' + ticketLink + '</a>' +
        '</p>' +

        '<hr style="border: 0; border-top: 1px solid #eeeeee; margin: 30px 0;">' +

        '<p style="color: #999999; font-size: 12px; text-align: center;">' +
        'Best regards,<br><strong>BIS AutoCAD Internship Team</strong>' +
        '</p>' +
        '</div>';

    try {
        MailApp.sendEmail({
            to: email,
            subject: subject,
            htmlBody: htmlBody
        });
        Logger.log("Email Sent Successfully to: " + email);
    } catch (e) {
        console.error("Email Failed to " + email + ": " + e.toString());
        Logger.log("Email Failed: " + e.toString());
    }
}

function getHeaders_(sheet) {
    return sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
}

function generateTicketId() {
    return 'BIS-ACAD-' + Math.random().toString(36).substr(2, 8).toUpperCase();
}

function generateToken() {
    return Utilities.getUuid().replace(/-/g, '').substring(0, 24);
}

function saveToDrive_(b64OrBlob, mimeType, fileName, folderName) {
    try {
        var folder = getOrCreateFolder_(folderName);
        var blob;
        if (typeof b64OrBlob === 'string') {
            var cleanB64 = b64OrBlob.includes("base64,") ? b64OrBlob.split("base64,")[1] : b64OrBlob;
            blob = Utilities.newBlob(Utilities.base64Decode(cleanB64), mimeType, fileName);
        } else {
            blob = b64OrBlob;
            blob.setName(fileName);
        }
        var file = folder.createFile(blob);
        // file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        return file.getUrl();
    } catch (e) {
        return "Error saving: " + e.toString();
    }
}

function saveToDriveBlob_(blob, folderName, fileName) {
    return saveToDrive_(blob, blob.getContentType(), fileName, folderName);
}

function getOrCreateFolder_(name) {
    var folders = DriveApp.getFoldersByName(name);
    return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function escapeHtml(text) {
    if (!text) return "";
    return text.toString().replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;").replace(/'/g, "&#039;");
}

function jsonResponse(data) {
    var output = ContentService.createTextOutput(JSON.stringify(data));
    output.setMimeType(ContentService.MimeType.JSON);
    return addCorsHeaders(output);
}

function addCorsHeaders(output) {
    return output;
}

function getMimeTypeFromFilename(filename) {
    if (filename.endsWith(".png")) return "image/png";
    if (filename.endsWith(".jpg") || filename.endsWith(".jpeg")) return "image/jpeg";
    return "application/octet-stream";
}

// --- PROXY FUNCTIONS ---

/**
 * Extracts file ID from various Google Drive URL formats.
 */
function fileIdFromDriveUrl(url) {
    if (!url) return null;
    var match = url.match(/\/d\/([a-zA-Z0-9_-]+)/);
    if (match) return match[1];

    match = url.match(/id=([a-zA-Z0-9_-]+)/);
    if (match) return match[1];

    // Fallback for raw ID (if it looks like an ID)
    if (url.match(/^[a-zA-Z0-9_-]{20,}$/)) return url;

    return null;
}

/**
 * Proxies a Drive image to base64 data URL to avoid CORS/403 issues.
 */
function handleProxyImage(params) {
    var fileId = params.fileId || params.id;

    if (!fileId) {
        return jsonResponse({ success: false, error: 'missing_fileId' });
    }

    try {
        var file = DriveApp.getFileById(fileId);
        var blob = file.getBlob();
        var b64 = Utilities.base64Encode(blob.getBytes());
        var mime = blob.getContentType();

        return jsonResponse({
            success: true,
            dataUrl: "data:" + mime + ";base64," + b64,
            mime: mime
        });

    } catch (e) {
        Logger.log("handleProxyImage Error: " + e.toString());
        return jsonResponse({ success: false, error: "Proxy failed: " + e.toString() });
    }
}
