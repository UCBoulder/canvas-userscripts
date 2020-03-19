// ==UserScript==
// @name         Export Rubric Scores
// @namespace    https://github.com/CUBoulder-OIT
// @description  Export all rubric criteria scores for an assignment to a CSV
// @include      https://canvas.*.edu/courses/*/gradebook/speed_grader?*
// @include      https://*.*instructure.com/courses/*/gradebook/speed_grader?*
// @grant        none
// @run-at       document-idle
// @version      1.0.0
// ==/UserScript==

/* globals $ */

// wait until the window jQuery is loaded
function defer(method) {
    if (typeof $ !== 'undefined') {
        method();
    }
    else {
        setTimeout(function() { defer(method); }, 100);
    }
}

function waitForElement(selector, callback) {
    if ($(selector).length) {
        callback();
    } else {
        setTimeout(function() {
            waitForElement(selector, callback);
        }, 100);
    }
}

function popUp(text) {
    $("#export_rubric_dialog").html(`<p>${text}</p>`);
    $("#export_rubric_dialog").dialog({ buttons: {} });
}

function popClose() {
    $("#export_rubric_dialog").dialog("close");
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

defer(function() {
    'use strict';

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

    $("body").append($('<div id="export_rubric_dialog" title="Export Rubric Scores"></div>'));
    // Only add the export button if a rubric is appearing
    if ($('#rubric_summary_holder').length > 0) {
        $('#gradebook_header div.statsMetric').append('<button type="button" class="Button" id="export_rubric_btn">Export Rubric Scores</button>');
        $('#export_rubric_btn').click(function() {
            popUp("Exporting scores, please wait...");

            // Get some initial data from the current URL
            var courseId = window.location.href.split('/')[4];
            var urlParams = window.location.href.split('?')[1].split('&');
            var assignId;
            for (const param of urlParams) {
                if (param.split('=')[0] === "assignment_id") {
                    assignId = param.split('=')[1];
                    break;
                }
            }

            // Get the rubric data
            $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`, function(assignment) {
                // Get the user data
                getAllPages(`/api/v1/courses/${courseId}/enrollments?per_page=100`, function(enrollments) {
                    function getUser(userId) {
                        var user = null;
                        $.each(enrollments, function(enrollIndex, enrollment) {
                            if (enrollment.user_id === userId) {
                                user = enrollment.user;
                            }
                        });
                        return user;
                    }

                    // Get the rubric score data
                    getAllPages(`/api/v1/courses/${courseId}/assignments/${assignId}/submissions?include[]=rubric_assessment&per_page=100`, function(submissions) {

                        // Fill out the csv header and map criterion ids to sort index
                        var critOrder = {};
                        var header = "Student Name,Student ID,Posted Score,Attempt Number";
                        $.each(assignment.rubric, function(critIndex, criterion) {
                            critOrder[criterion.id] = critIndex;
                            header += `,Points: ${criterion.description}`;
                        });
                        header += '\n';

                        var csvRows = [header];
                        var subCount = 0;

                        // Function to call for each user/submission, to output data when finished
                        function recordSubmission(csvRow) {
                            subCount++;
                            csvRows.push(csvRow);
                            if (subCount >= submissions.length) {
                                popClose();
                                saveText(csvRows, `Rubric Scores ${assignment.name.replace(/[^a-zA-Z 0-9]+/g)}.csv`);
                            }
                        }

                        // Iterate through submissions
                        $.each(submissions, function(subIndex, submission) {
                            var user = getUser(submission.user_id);
                            if (user != null) {
                                var row = `${user.name},${user.sis_user_id},${submission.score},${submission.attempt}`;
                                // Add criteria scores
                                // Need to turn rubric_assessment object into an array
                                var crits = []
                                var critIds = []
                                if (submission.rubric_assessment != null) {
                                    $.each(submission.rubric_assessment, function(critKey, critValue) {
                                        crits.push({'id': critKey, 'points':critValue.points});
                                        critIds.push(critKey);
                                    });
                                }
                                // Check for any scores that might be missing; set them to null
                                $.each(critOrder, function(critKey, critValue) {
                                    if (!critIds.includes(critKey)) {
                                        crits.push({'id': critKey, 'points':null});
                                    }
                                });
                                // Sort into same order as column order
                                crits.sort(function(a, b) { return critOrder[a.id] - critOrder[b.id]; });
                                $.each(crits, function(critIndex, criterion) {
                                    row += `,${criterion.points}`;
                                });
                                row += '\n';
                                recordSubmission(row);
                            } else {
                                // Still need to record something so that we'll know when all submissions have been checked
                                recordSubmission(`Error: Could not find user ${submission.user_id}\n`);
                            }
                        });
                    });
                });
            });
        });
    }
});
