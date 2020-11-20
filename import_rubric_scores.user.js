// ==UserScript==
// @name         Import Rubric Scores
// @namespace    https://github.com/CUBoulder-OIT
// @description  Import rubric criteria scores or ratings for an assignment from a CSV
// @include      https://canvas.*.edu/courses/*/gradebook/speed_grader?*
// @include      https://*.*instructure.com/courses/*/gradebook/speed_grader?*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.1.0/papaparse.min.js
// @run-at       document-idle
// @version      0.1.0
// ==/UserScript==

/* globals $ Papa */

// wait until the window jQuery is loaded
function defer(method) {
    if (typeof $ !== 'undefined') {
        method();
    }
    else {
        setTimeout(function() { defer(method); }, 100);
    }
}

// utility function for downloading a file
var saveText = (function () {
    var a = document.createElement("a");
    document.body.appendChild(a);
    a.style = "display: none";
    return function (textArray, fileName) {
        var blob = new Blob(textArray, {type: "text"}),
            url = window.URL.createObjectURL(blob);
        a.href = url;
        a.download = fileName;
        a.click();
        window.URL.revokeObjectURL(url);
    };
}());

// Pop up that calls "callback" when Ok is pressed
// If callback is null, then it isn't used
function popUp(text, callback) {
    $("#import_rubric_popup_dialog").html(`<p>${text}</p>`);
    if (callback !== null) {
        $("#import_rubric_popup_dialog").dialog({ buttons: { Ok: function() { $(this).dialog("close"); callback(); } } });
    } else {
        $("#import_rubric_popup_dialog").dialog({ buttons: { Ok: function() { $(this).dialog("close"); } } });
    }
    $("#import_rubric_popup_dialog").dialog("open");
}

function showProgress(amount) {
    if (amount === 100) {
        $("#import_rubric_progress").dialog("close");
    } else {
        $("#import_rubric_bar").progressbar({ value: amount });
        $("#import_rubric_progress").dialog("open");
    }
}

function getAllPages(url, callback) {
    getRemainingPages(url, [], callback);
}

// Recursively work through paginated JSON list
function getRemainingPages(nextUrl, listSoFar, callback) {
    $.getJSON(nextUrl, function(responseList, textStatus, jqXHR) {
        var nextLink = null;
        $.each(jqXHR.getResponseHeader("link").split(','), function (linkIndex, linkEntry) {
            if (linkEntry.split(';')[1].includes('rel="next"')) {
                nextLink = linkEntry.split(';')[0].slice(1, -1);
            }
        });
        if (nextLink == null) {
            // all pages have been retrieved
            callback(listSoFar.concat(responseList));
        } else {
            getRemainingPages(nextLink, listSoFar.concat(responseList), callback);
        }
    });
}

// Send a list of requests in chunks of 10 to avoid rate limiting
function sendRequests(requests, successCallback, errorCallback) {
    var errors = [];
    var completed = 0;
    var chunkSize = 10;
    function sendChunk(i) {
        for (const request of requests.slice(i, i + chunkSize)) {
            $.ajax(request.request).fail(function(jqXHR, textStatus, errorThrown) {
                errors.push(`${request.error}${jqXHR.status} - ${errorThrown}\n`);
            }).always(requestSent);
        }
        showProgress(i * 100 / requests.length);
        if (i + chunkSize < requests.length) {
            setTimeout(sendChunk, 1000, i + chunkSize);
        }
    }

    // when each request finishes...
    function requestSent() {
        completed++;
        if (completed >= requests.length) {
            // all finished
            showProgress(100);
            if (errors.length > 0) {
                errorCallback(errors);
            } else {
                successCallback();
            }
        }
    }
    // actually starts the recursion
    sendChunk(0);
}

// Opens the main UI.
// fileCallback should accept the file to validate along with a success callback that will be passed the csvData
// importCallback should accept the csv data and will actually import it
function openImportDialog(fileCallback, importCallback) {
    // build HTML
    $("#import_rubric_dialog").html(`
<p>Import a spreadsheet of rubric scores or ratings as grades for this assignment.</p>
<hr>
<label for="import_rubric_file">Rubric scores file: </label>
<input type="file" id="import_rubric_file"/><br>
<span id="import_rubric_results"></span><br>
<button type="button" class="Button" id="import_rubric_import_btn" disabled>Import</button>`);

    // when file is selected, validate data and enable Import button
    $('#import_rubric_file').change(function(evt) {
        fileCallback(evt.target.files[0], function(csvData) {
            $('#import_rubric_import_btn').removeAttr("disabled");
            $('#import_rubric_import_btn').click(function() { $("#import_rubric_dialog").dialog("close"); importCallback(csvData); });
        });
    });
    $("#import_rubric_dialog").dialog({width: "375px"});
}

// Check that the spreadsheet data is in the correct format
function validateData(csvData, successCallback) {
    $("#import_rubric_results").text(`Found ${csvData.length} rows.`);
    successCallback(csvData);
}

// Actually import the data user by user
// arg is a list of dicts with members:
//  - user_id: the Canvas user id
//  - criterion_{id}: the points that should be given for this criterion
function importScores(csvData) {

    // Get some initial data from the current URL
    const courseId = window.location.href.split('/')[4];
    const urlParams = window.location.href.split('?')[1].split('&');
    var assignId;
    for (const param of urlParams) {
        if (param.split('=')[0] === "assignment_id") {
            assignId = param.split('=')[1];
            break;
        }
    }

    // Get assignment data
    $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`, function(assignment) {

        // Build requests
        var requests = [];
        for (const row of csvData) {
            const endpoint = `/api/v1/courses/${courseId}/assignments/${assignId}/submissions/sis_user_id:${row["Student ID"]}`;
            var params = {};
            for (const prop in row) {
                if (prop.startsWith("Points: ")) {
                    const criterionId = assignment.rubric.find( c => c.description === prop.slice(8) ).id;
                    params[`rubric_assessment[${criterionId}][points]`] = row[prop];
                }
            }
            requests.push({request: {url: endpoint,
                                     type: "PUT",
                                     data: params,
                                     dataType: "text" },
                           error: `Failed to import scores for student ${row["Student ID"]} using endpoint ${endpoint}. Response: `});
        }
        sendRequests(
            requests,
            function() { popUp("All scores/ratings imported successfully!", function() { location.reload(); }); },
            function(errors) {
                saveText(errors, "errors.txt");
                popUp(`Import complete. WARNING: ${errors.length} rows failed to import. See errors.txt for details.`, function() { location.reload(); });
            });
    });
}

defer(function() {
    'use strict';

    $("body").append($('<div id="import_rubric_popup_dialog" title="Import Rubric Scores"></div>'));
    $("body").append($('<div id="import_rubric_dialog" title="Import Rubric Scores"></div>'));
    $("body").append($('<div id="import_rubric_progress" title="Import Rubric Scores"><p>Importing rubric scores. Do not navigate from this page.</p><div id="import_rubric_bar"></div></div>'));
    $("#import_rubric_progress").dialog({ buttons: {}, autoOpen: false });

    // Only add the import button if a rubric is appearing
    if ($('#rubric_summary_holder').length > 0) {
        $('#gradebook_header div.statsMetric').append('<button type="button" class="Button" id="import_rubric_btn">Import Rubric Scores</button>');
        $('#import_rubric_btn').click(function() {
            openImportDialog(function(importFile, successCallback) {
                // parse CSV
                Papa.parse(importFile, {
                    header: true,
                    dynamicTyping: false,
                    complete: function(results) {
                        $("#import_rubric_file").val('');
                        validateData(results.data, successCallback);
                    }
                });
            }, importScores);
        });
    }
});
