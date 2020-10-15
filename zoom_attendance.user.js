// ==UserScript==
// @name         Import Zoom Attendance
// @namespace    https://github.com/CUBoulder-OIT
// @description  Create a graded attendance assignment based on a Zoom participants export.
// @include      https://canvas.*.edu/*/gradebook
// @include      https://*.*instructure.com/*/gradebook
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.1.0/papaparse.min.js
// @run-at       document-idle
// @version      1.3.0
// ==/UserScript==

/* globals $ Papa */

// wait until the window jQuery is loaded
function defer(method) {
    if (typeof $ !== 'undefined' && typeof $().dialog !== 'undefined') {
        method();
    }
    else {
        setTimeout(function() { defer(method); }, 100);
    }
}

// utility function for downloading an error report
var saveText = (function () {
    let a = document.createElement("a");
    document.body.appendChild(a);
    a.style = "display: none";
    return function (textArray, fileName) {
        let blob = new Blob(textArray, {type: "text"}),
            url = window.URL.createObjectURL(blob);
        a.href = url;
        a.download = fileName;
        a.click();
        window.URL.revokeObjectURL(url);
    };
}());

// utility function for getting value based on field name from serialized form data
function getFormValue(name, serial) {
    var dc = s => decodeURIComponent(s.replace(/\+/g, '%20'));
    for (const elem of serial.split('&')) {
        var pair = elem.split('=');
        if (dc(pair[0]) === name) {
            return dc(pair[1]);
        }
    }
}

function getAllPages(url, callback) {
    getRemainingPages(url, [], callback);
}

// Recursively work through paginated JSON list
function getRemainingPages(nextUrl, listSoFar, callback) {
    $.getJSON(nextUrl, function(responseList, textStatus, jqXHR) {
        let nextLink = null;
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
    }).fail(function (jqXHR, textStatus, errorThrown) {
        popUp(`ERROR ${jqXHR.status} while retrieving data from Canvas. Please refresh and try again.`, null);
        $("#zoom_file").show();
    });
}

function openOptionsForm(assignName, submitAction) {
    // built select options from assignment groups list
    let courseId = window.location.href.split('/')[4];
    $.getJSON(`/api/v1/courses/${courseId}/assignment_groups?per_page=100`, function(assignGroups) {
        let groupOptions = '';
        $.each(assignGroups, function(index, group) {
            groupOptions = `${groupOptions}
<option value="${group.id}">${group.name}</option>`;
        });

        // build HTML
        $("#zoom_form_dialog").html(`
<p>
Import a Zoom Participants spreadsheet as attendance grades to a new column (assignment).
</p>
<hr>
<form id="zoom_form">
<label>New assignment name</label>
<input autocomplete="off" name="zoom_assign_name" value="${assignName}">
<br>
<br>
<label>Assignment points</label>
<input type="number" autocomplete="off" name="zoom_points" value=1 style="width:5em">
<br>
<br>
<label>Assignment group</label>
<select name="zoom_assign_group">
${groupOptions}
</select>
<br>
<br>
<label>Minimum minutes to be counted present</label>
<input type="number" autocomplete="off" name="zoom_min_minutes" value=1 style="width:5em">
<br>
<br>
<input type="submit" value="Import" class="btn btn-primary">
</form>`);

        // on submit, send form data to submitAction
        $("#zoom_form").submit(function(event) {
            event.preventDefault();
            $("#zoom_form_dialog").dialog("close");
            submitAction($("#zoom_form").serialize());
        });
        $("#zoom_form_dialog").on("dialogclose", function () {
            $("#zoom_file").show();
        });
        $("#zoom_form_dialog").dialog("open");

    }).fail(function (jqXHR, textStatus, errorThrown) {
        popUp(`ERROR ${jqXHR.status} while getting list of assignment groups. Please refresh and try again.`, null);
        $("#zoom_file").show();
    });
}

function showProgress(amount) {
    if (amount === 100) {
        $("#zoom_progress").dialog("close");
    } else {
        $("#zoom_bar").progressbar({ value: amount });
        $("#zoom_progress").dialog("open");
    }
}

// Pop up that calls "callback" when Ok is pressed
// If callback is null, then it isn't used
function popUp(text, callback) {
    $("#zoom_dialog").html(`<p>${text}</p>`);
    if (callback !== null) {
        $("#zoom_dialog").dialog({ buttons: { Ok: function() { $(this).dialog("close"); callback(); } } });
    } else {
        $("#zoom_dialog").dialog({ buttons: { Ok: function() { $(this).dialog("close"); } } });
    }
    $("#zoom_dialog").dialog("open");
}

// Open pop up if condition is true, otherwise go straight to callback
function popUpIf(condition, text, callback) {
    if (condition) {
        popUp(text, callback);
    } else {
        callback();
    }
}

// Turn an Excel-style serialDate into a date string
function dateFromSerial(serialDate) {
    const date = new Date(Date.UTC(0, 0, serialDate, 0));
    return `${date.getMonth() + 1}-${date.getDate()}-${date.getFullYear()}`
}

// Get the data out of the selected spreadsheet file
function parseImport(importJson, callback) {

    // Validate spreadsheet format
    let headerRow = null;
    for (let i=0; i<importJson.length; i++) {
        if (importJson[i].length > 0 && importJson[i][0].toLowerCase().includes('name')) {
            headerRow = i;
            break;
        }
    }
    if (headerRow === null) {
        popUp(`ERROR - Could not identify the name column of the spreadsheet. Re-download the Zoom participants export and try again.`, null);
        $("#zoom_file").show();
        return;
    }
    if (!(importJson[headerRow].length > 1 && importJson[headerRow][1].toLowerCase().includes('email'))) {
        popUp(`ERROR - Could not identify the email column of the spreadsheet. Re-download the Zoom participants export and try again.`, null);
        $("#zoom_file").show();
        return;
    }
    let minutesCol = null;
    for (let i=2; i<importJson[headerRow].length; i++) {
        if (importJson[headerRow][i].toLowerCase().includes('duration')) {
            minutesCol = i;
            break;
        }
    }
    if (minutesCol === null) {
        popUp(`ERROR - Could not identify the duration column of the spreadsheet. Re-download the Zoom participants export and try again.`, null);
        $("#zoom_file").show();
        return;
    }
    let startTimeCol = null;
    for (let i=0; i<importJson[0].length; i++) {
        if (importJson[0][i].toLowerCase().includes('start time')) {
            startTimeCol = i;
            break;
        }
    }
    const startDate = startTimeCol !== null && importJson[1].length > startTimeCol ? importJson[1][startTimeCol].slice(0, 10) : "Unknown Date";

    // Extract users from the file
    let importUsers = [];
    let noImports = [];
    for (const row of importJson.slice(headerRow + 1)) {
        if (row.length >= 1 && row[1] && row[1].includes('@')) {
            const minutes = row.length >= 2 ? row[minutesCol] : 0;
            importUsers.push({'username': row[1].split("@")[0], 'matched': false, 'minutes': minutes});
        } else if (row.length > 0 && row[0]) {
            noImports.push(row[0]);
        }
    }
    const customText = noImports.length > 10 ? "The first ten of these u" : "U";
    popUpIf(noImports.length > 0, `NOTICE - ${noImports.length} users will not be imported because their email was not captured by Zoom. ${customText}sers are listed below.<br><br>${noImports.slice(0, 10).join('<br>')}`, function() {

        // Next, get a list of active students in the course
        let courseId = window.location.href.split('/')[4];
        popUp(`Checking course roster. Please wait...`, null);
        getAllPages(`/api/v1/courses/${courseId}/enrollments?type[]=StudentEnrollment&state[]=active&per_page=100`, function(enrollments) {
            let attendData = { 'date': startDate, 'users': [] };
            // For each student in the course, note the minutes if they're in the import; otherwise give them 0 minutes
            // Note which of the import users we were able to find matches in the course for
            for (let i=0; i<enrollments.length; i++) {
                let match = importUsers.find(item => item.username == enrollments[i].user.login_id);
                if (match) {
                    match.matched = true;
                    attendData.users.push({ 'userId': enrollments[i].user_id, 'username': match.username, 'minutes': parseFloat(match.minutes) });
                } else {
                    attendData.users.push({ 'userId': enrollments[i].user_id, 'username': enrollments[i].user.login_id, 'minutes': 0 });
                }
            }
            // Build a list of unmatched students to notify about
            let unmatched = [];
            for (const item of importUsers.filter(item => !item.matched)) {
                unmatched.push(item.username);
            }
            const customText = noImports.length > 10 ? "The first ten of these u" : "U";
            popUpIf(unmatched.length > 0, `NOTICE - ${unmatched.length} users will not be imported because their username does not match an enrolled Canvas user. ${customText}sers are listed below.<br><br>${unmatched.slice(0, 10).join('<br>')}`, function() {
                // Close the "Please wait" popup if it's still open
                $('#zoom_dialog').dialog("close");
                callback(attendData);
            });
        });
    });
}

// Actually save the grades to Canvas based on data from the spreadsheet and the dialog form
function saveAttendance(formData, attendData) {
    const assignName = getFormValue('zoom_assign_name', formData);
    const assignPoints = getFormValue('zoom_points', formData);
    const minMinutes = getFormValue('zoom_min_minutes', formData);
    const assignGroup = getFormValue('zoom_assign_group', formData);
    const courseId = window.location.href.split('/')[4];

    // create assignment
    $.ajax({
        url: `/api/v1/courses/${courseId}/assignments`,
        type: 'POST',
        data: { 'assignment[name]': assignName,
               'assignment[submission_types][]': 'none',
               'assignment[points_possible]': assignPoints,
               'assignment[assignment_group_id]': assignGroup,
               'assignment[published]': true },
        dataType: "text"
    }).fail(function (jqXHR, textStatus, errorThrown) {
        popUp(`ERROR ${jqXHR.status} while creating assignment. Please refresh and try again.`, null);
        $("#zoom_file").show();
    }).done(function (responseText) {

        // "mute" assignment (i.e. set the post policy to "manual" so students won't see the grades by default)
        const assignId = JSON.parse(responseText).id;
        $.ajax({
            url: `/api/graphql`,
            type: 'POST',
            data: String.raw`{"operationName":"SetAssignmentPostPolicy","variables":{"assignmentId":"${assignId}","postManually":true},"query":"mutation SetAssignmentPostPolicy($assignmentId: ID!, $postManually: Boolean!) {\n  setAssignmentPostPolicy(input: {assignmentId: $assignmentId, postManually: $postManually}) {\n    postPolicy {\n      postManually\n      __typename\n    }\n    errors {\n      attribute\n      message\n      __typename\n    }\n    __typename\n  }\n}\n"}`,
            contentType: "application/json",
            dataType: "text"
        }).fail(function (jqXHR, textStatus, errorThrown) {
            popUp(`ERROR ${jqXHR.status} while muting assignment. Please delete the empty assignment before trying again.`, null);
            $("#zoom_file").show();
        }).done(function (responseText2) {

            // prepare a list of requests to send
            let requests = [];
            for (const user of attendData.users) {
                // build api url
                const subUrl = `/api/v1/courses/${courseId}/assignments/${assignId}/submissions/${user.userId}`;
                // build request and canned error message in case it fails
                requests.push({
                    request: {
                        url: subUrl,
                        type: "PUT",
                        data: { 'submission[posted_grade]': user.minutes >= minMinutes ? assignPoints : 0 },
                        dataType: "text" },
                    error: `Failed to post score for student ${user.username} (id: ${user.userId}) using endpoint ${subUrl}. Response: `
                });
            }

            // send requests in chunks of 10 every second to avoid rate-limiting
            var errors = [];
            var completed = 0;
            var chunkSize = 10;
            function sendChunk(i) {
                for (const request of requests.slice(i, i + chunkSize)) {
                    $.ajax(request.request).fail(function(jqXHR, textStatus, errorThrown) {
                        if (jqXHR.status == 500) {
                            // Canvas sometimes gets random server errors, so retry
                            $.ajax(request.request).fail(function(jqXHR, textStatus, errorThrown) {
                                errors.push(`${request.error}${jqXHR.status} - ${errorThrown}\n`);
                            });
                        } else {
                            errors.push(`${request.error}${jqXHR.status} - ${errorThrown}\n`);
                        }
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
                    $("#zoom_file").show();
                    if (errors.length > 0) {
                        popUp(`Import complete. WARNING: ${errors.length} scores failed to import. See errors.txt for details and then manually enter the correct scores for these users.`, function() { window.location.href = window.location.href; });
                        saveText(errors, "errors.txt");
                    } else {
                        popUp("All scores imported successfully!", function() { window.location.href = window.location.href; });
                    }
                }
            }
            // actually starts the recursion
            sendChunk(0);
        });

    });

}

defer(function() {
    'use strict';

    // prep jquery form dialog
    $("body").append($('<div id="zoom_form_dialog" title="Import Zoom Attendance"></div>'));
    $("#zoom_form_dialog").dialog({ autoOpen: false, width: "30em" });

    // prep jquery info dialog
    $("head").append($('<style type="text/css"> .no-close .ui-dialog-titlebar-close { display: none; } </style>'))
    $("body").append($('<div id="zoom_dialog" title="Import Zoom Attendance"></div>'));
    $("#zoom_dialog").dialog({ autoOpen: false, dialogClass: "no-close", closeOnEscape: false });

    // prep jquery progress dialog
    $("body").append($('<div id="zoom_progress" title="Import Zoom Attendance"><p>Importing attendance. Do not navigate from this page.</p><div id="zoom_bar"></div></div>'));
    $("#zoom_progress").dialog({ autoOpen: false, buttons: {} });

    // add choose file button to gradebook
    let importDiv = $(`<div style="padding-top:10px;>
<label for="zoom_file">Import Zoom attendance: </label>
<input type="file" id="zoom_file"/>
</div>`);
    $("div.gradebook-menus").append(importDiv);

    // handle when file is selected
    $('#zoom_file').change(function(evt) {
        if (!evt.target.files[0].name.endsWith('.csv')) {
            popUp(`Must import a .csv spreadsheet downloaded from Zoom.`, null);
        } else {
            Papa.parse(evt.target.files[0], {
                complete: function(results) {
                    parseImport(results.data, function(attendanceData) {
                        openOptionsForm(attendanceData.date, function(formData) {
                            saveAttendance(formData, attendanceData);
                        });
                    });}
            });
            $("#zoom_file").hide();
        }
        $("#zoom_file").val('');
    });
});
