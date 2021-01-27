// ==UserScript==
// @name         Export Rubric Scores
// @namespace    https://github.com/CUBoulder-OIT
// @description  Export all rubric criteria scores for an assignment to a CSV
// @include      https://canvas.*.edu/courses/*/gradebook/speed_grader?*
// @include      https://*.*instructure.com/courses/*/gradebook/speed_grader?*
// @grant        none
// @run-at       document-idle
// @version      1.1.0
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
            const courseId = window.location.href.split('/')[4];
            const urlParams = window.location.href.split('?')[1].split('&');
            const assignId = urlParams.find(i => i.split('=')[0] === "assignment_id").split('=')[1];

            // Get the rubric data
            $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`, function(assignment) {
                // Get the user data
                getAllPages(`/api/v1/courses/${courseId}/enrollments?per_page=100`, function(enrollments) {
                    // Get the rubric score data
                    getAllPages(`/api/v1/courses/${courseId}/assignments/${assignId}/submissions?include[]=rubric_assessment&per_page=100`, function(submissions) {
                        // If rubric is set to hide points, then also hide points in export
                        // If rubric is set to use free form comments, then also hide ratings in export
                        const hidePoints = assignment.rubric_settings.hide_points;
                        const hideRatings = assignment.rubric_settings.free_form_criterion_comments;
                        if (hidePoints && hideRatings) {
                            popUp("ERROR: This rubric is configured to use free-form comments instead of ratings AND to hide points, so there is nothing to export!");
                            return;
                        }

                        // Fill out the csv header and map criterion ids to sort index
                        // Also create an object that maps criterion ids to an object mapping rating ids to descriptions
                        var critOrder = {};
                        var critRatingDescs = {};
                        var header = "Student Name,Student ID,Posted Score,Attempt Number";
                        $.each(assignment.rubric, function(critIndex, criterion) {
                            critOrder[criterion.id] = critIndex;
                            critRatingDescs[criterion.id] = {};
                            $.each(criterion.ratings, function(i, rating) {
                                critRatingDescs[criterion.id][rating.id] = rating.description;
                            });
                            if (!hideRatings) {
                                header += `,Rating: ${criterion.description}`;
                            }
                            if (!hidePoints) {
                                header += `,Points: ${criterion.description}`;
                            }
                        });
                        header += '\n';

                        // Iterate through submissions
                        var csvRows = [header];
                        $.each(submissions, function(subIndex, submission) {
                            const user = enrollments.find(i => i.user_id === submission.user_id).user;
                            if (user) {
                                var row = `${user.name},${user.sis_user_id},${submission.score},${submission.attempt}`;
                                // Add criteria scores and ratings
                                // Need to turn rubric_assessment object into an array
                                var crits = []
                                var critIds = []
                                if (submission.rubric_assessment != null) {
                                    $.each(submission.rubric_assessment, function(critKey, critValue) {
                                        crits.push({'id': critKey, 'points': critValue.points, 'rating': critRatingDescs[critKey][critValue.rating_id]});
                                        critIds.push(critKey);
                                    });
                                }
                                // Check for any criteria entries that might be missing; set them to null
                                $.each(critOrder, function(critKey, critValue) {
                                    if (!critIds.includes(critKey)) {
                                        crits.push({'id': critKey, 'points': null, 'rating': null});
                                    }
                                });
                                // Sort into same order as column order
                                crits.sort(function(a, b) { return critOrder[a.id] - critOrder[b.id]; });
                                $.each(crits, function(critIndex, criterion) {
                                    if (!hideRatings) {
                                        row += `,${criterion.rating}`
                                    }
                                    if (!hidePoints) {
                                        row += `,${criterion.points}`;
                                    }
                                });
                                row += '\n';
                                csvRows.push(row);
                            }
                        });
                        popClose();
                        saveText(csvRows, `Rubric Scores ${assignment.name.replace(/[^a-zA-Z 0-9]+/g, '')}.csv`);
                    });
                });
            });
        });
    }
});
