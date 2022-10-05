// ==UserScript==
// @name         Import Rubric Scores
// @namespace    https://github.com/CUBoulder-OIT
// @description  Import rubric criteria scores or ratings for an assignment from a CSV
// @match        https://*/courses/*/gradebook/speed_grader?*
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.1.0/papaparse.min.js
// @run-at       document-idle
// @version      1.0.3
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
    $("#import_rubric_bar").progressbar({ value: amount });
    $("#import_rubric_progress").dialog();
    if (amount === 100) {
        $("#import_rubric_progress").dialog("close");
    } else {
        $("#import_rubric_progress").dialog("open");
    }
}

function getAssignId() {
    const urlParams = window.location.href.split('?')[1].split('&');
    for (const param of urlParams) {
        if (param.split('=')[0] === "assignment_id") {
            return param.split('=')[1];
        }
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
// requests is a list of objects with properties:
//  - request: The object for the jquery request
//  - error: The error message to display if the request fails
function sendRequests(requests, successCallback, errorCallback) {
    var errors = [];
    var completed = 0;
    var chunkSize = 10;
    function sendChunk(i) {
        $.each(requests.slice(i, i + chunkSize), function(i, request) {
            $.ajax(request.request).fail(function(jqXHR, textStatus, errorThrown) {
                errors.push(`${request.error}${jqXHR.status} - ${errorThrown}\n`);
            }).always(requestSent);
        });
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
// fileCallback should accept the file to validate along with a success callback that will be passed the import requests
// importCallback should accept the list of score objects and actually perform the import
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
        fileCallback(evt.target.files[0], function(scoreData) {
            $('#import_rubric_import_btn').removeAttr("disabled");
            $('#import_rubric_import_btn').click(function() { $("#import_rubric_dialog").dialog("close"); importCallback(scoreData); });
        });
    });
    $("#import_rubric_dialog").dialog({width: "375px"});
}

// Check that the spreadsheet data is in the correct format
// Passes a list of objects with score info to be imported to successCallback
function validateData(csvData, successCallback) {
    let inData = csvData.filter( i => i["Student ID"] !== undefined );
    let outData = inData.filter( i => i["Student ID"] !== "null" );
    let allPtsCriteria = Object.keys(outData[0]).filter( i => i.startsWith("Points: ")).map(i => i.slice(8));
    let allRtgCriteria = Object.keys(outData[0]).filter( i => i.startsWith("Rating: ")).map(i => i.slice(8));

    // Check for errors
    if (!inData.length) {
        $("#import_rubric_results").text(`No data in CSV. Please choose a different file.`);
    } if (!outData.length) {
        $("#import_rubric_results").text(`No users found with valid (not "null") Student IDs. Please choose a different file.`);
    } else if (!('Student Name' in outData[0])) {
        $("#import_rubric_results").text(`No "Student Name" column found. Please double-check file format.`);
    } else if (!('Student ID' in outData[0])) {
        $("#import_rubric_results").text(`No "Student ID" column found. Please double-check file format.`);
    } else if (!allPtsCriteria.length && !allRtgCriteria.length) {
        $("#import_rubric_results").text(`No column headers found in the form of "Points: {criterion name}" or "Rating: {criterion name}". Please double-check file format.`);
    } else {
        const courseId = window.location.href.split('/')[4];
        const assignId = getAssignId();

        // Get assignment data to match up criteria from file with those in Canvas
        $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`, function(assignment) {
            // Identify ids to go with the criteria and build a list of warnings while we're at it
            let criteriaIds = {};
            let unmatchedCriteria = [];
            let doubledCriteria = [];
            $.each(allPtsCriteria, function(i, criterion) {
                const id = assignment.rubric.find(i => i.description === criterion).id;
                if (id) {
                    criteriaIds[criterion] = id;
                } else {
                    unmatchedCriteria.push(criterion);
                }
            });
            $.each(allRtgCriteria, function(i, criterion) {
                const id = assignment.rubric.find(i => i.description === criterion).id;
                if (id) {
                    if (criterion in criteriaIds) {
                        // Giving a criteria a rating when it already has a points value is redundant at best
                        doubledCriteria.push(criterion);
                    } else {
                        criteriaIds[criterion] = id;
                    }
                } else {
                    unmatchedCriteria.push(criterion);
                }
            });
            if (!Object.keys(criteriaIds).length) {
                $("#import_rubric_results").text(`No criteria listed in the file match those listed for this assignment's rubric in Canvas. Please double-check file format.`);
            } else {

                // Build list of score objects
                function getRatingPoints(criterionId, ratingDesc) {
                    const criterion = assignment.rubric.find(i => i.id === criterionId);
                    return criterion.ratings.find(i => i.description === ratingDesc).points;
                }
                function getRatingObj(criterionId, ratingDesc) {
                    const criterion = assignment.rubric.find(i => i.id === criterionId);
                    return criterion.ratings.find(i => i.description === ratingDesc);
                }
                function getScores(row) {
                    let scores = {user: row['Student ID']};
                    $.each(criteriaIds, function(criterion, critId) {
                        // also map criteria ids each to an object giving the 'points' and possibly also the 'rating'
                        // You can't just give Canvas a rating, even though Canvas knows how many points it's worth: you must always specify points
                        if (`Points: ${criterion}` in row) {
                            // When points are specified, ignore the rating, which might contradict the points and cause an error
                            scores[critId] = {points: row[`Points: ${criterion}`]};
                        } else {
                            // When the rating alone is specified, we still need to dig up how many points it should be worth, or it will error
                            const rating = getRatingObj(critId, row[`Rating: ${criterion}`]);
                            scores[critId] = rating ? {points: rating.points, rating: rating.id} : {error: String(row[`Rating: ${criterion}`])};
                        }
                    });
                    return scores;
                }
                const scoreData = outData.map(i => getScores(i));
                const badRows = scoreData.filter(i => Object.keys(i).find(j => j !== 'user' && 'error' in i[j] && i[j].error !== 'undefined' && i[j].error !== ''));
                if (badRows.length) {
                    let errors = [];
                    $.each(badRows.slice(0, 5), function(i, scoreObj) {
                        const badRating = scoreObj[Object.keys(scoreObj).find(i => i !== 'user' && 'error' in scoreObj[i])].error;
                        errors.push(`ERROR: Rating "${badRating}" given for user ${scoreObj.user} does not match a rating on this rubric in Canvas.`);
                    });
                    let errTxt = errors.join('<br>');
                    if (errors.length < badRows.length) {
                        errTxt += `<br>...<br>(${badRows.length - errors.length} additional rating errors not shown)`;
                    }
                    $("#import_rubric_results").html(errTxt);
                } else {
                    // File is runnable, now display any warnings
                    let notice = `<p>Ready to import scores for ${Object.keys(criteriaIds).length} criteria and ${outData.length} user(s).</p>`;
                    if (unmatchedCriteria.length) {
                        notice += `<p>WARNING: These ${unmatchedCriteria.length} criteria could not be found in this assignment's rubric in Canvas and will be ignored:<br>${unmatchedCriteria.join('<br>')}.</p>`;
                    }
                    if (doubledCriteria.length) {
                        notice += `<p>WARNING: These ${doubledCriteria.length} criteria have both a rating and a points value given; the ratings will be ignored and auto-assigned by Canvas:<br>${doubledCriteria.join('<br>')}.</p>`;
                    }
                    if ('Posted Score' in outData[0]) {
                        notice += `<p>Note: "Posted Score" column will be ignored.</p>`;
                    }
                    if (outData.length < inData.length) {
                        notice += `<p>Note: ${inData.length - outData.length} user(s) with a "null" Student ID will be ignored.</p>`;
                    }
                    $("#import_rubric_results").html(notice);
                    successCallback(scoreData);
                }
            }
        });
    }
}

// Actually import the data user by user
// scores is a list of objects with properties:
//  - user: The SIS User ID
//  - plus any number of criteria ids mapped each to an object with property 'points' and optionally 'rating'
function importScores(scores) {
    $("#import_rubric_file").val('');
    const courseId = window.location.href.split('/')[4];
    const assignId = getAssignId();

    // When each request is prepared, add it to the list
    // When the list is complete, send them off
    var requests = [];
    const total = scores.length;
    function pushRequest(request) {
        requests.push(request);
        if (requests.length === total) {
            sendRequests(
                requests,
                function() { popUp("All scores/ratings imported successfully!", function() { location.reload(); }); },
                function(errors) {
                    saveText(errors, "errors.txt");
                    popUp(`Import complete. WARNING: ${errors.length} rows failed to import. See errors.txt for details.`, function() { location.reload(); });
                });
        }
    }

    // Build each request by recursively iterating through scores to ensure GETs are sent one chunk at a time
    const chunkSize = 10;
    function buildRequests(chunkIndex) {
        $.each(scores.slice(chunkIndex, chunkIndex + chunkSize), function(index, userScore) {
            const endpoint = `/api/v1/courses/${courseId}/assignments/${assignId}/submissions/sis_user_id:${userScore.user}`;
            // Get existing rubric assessment from the submissions API
            // This allows us to ensure that existing data like comments aren't overwritten
            $.getJSON(`${endpoint}?include[]=rubric_assessment`, function(submission) {
                // Pre-load params with existing rubric assessment data
                var params = {};
                if (submission.rubric_assessment) {
                    $.each(submission.rubric_assessment, function(rowKey, rowValue) {
                        $.each(rowValue, function(cellKey, cellValue) {
                            params[`rubric_assessment[${rowKey}][${cellKey}]`] = cellValue;
                        });
                        // Make sure the comments param is never left out or undefined; Canvas can't handle this
                        if (!(`rubric_assessment[${rowKey}][comments]` in params) || params[`rubric_assessment[${rowKey}][comments]`] === undefined) {
                            params[`rubric_assessment[${rowKey}][comments]`] = "";
                        }
                    });
                }
                // Now fill in our points or ratings to be applied
                $.each(userScore, function(critId, critScore) {
                    if (critId !== 'user') {
                        if ('error' in critScore) {
                            // These are "undefined" ratings, in which case we'll just clear the rating and score;
                            params[`rubric_assessment[${critId}][points]`] = undefined;
                            params[`rubric_assessment[${critId}][rating_id]`] = undefined;
                        } else {
                            params[`rubric_assessment[${critId}][points]`] = critScore.points;
                            if ('rating' in critScore) {
                                params[`rubric_assessment[${critId}][rating_id]`] = critScore.rating;
                            } else {
                                delete params[`rubric_assessment[${critId}][rating_id]`];
                            }
                        }
                        // Again, ensure we don't leave out the comments field
                        if (!(`rubric_assessment[${critId}][comments]` in params)) {
                            params[`rubric_assessment[${critId}][comments]`] = "";
                        }
                    }
                });
                pushRequest({request: {url: endpoint,
                                       type: "PUT",
                                       data: params,
                                       dataType: "text" },
                             error: `Failed to import scores for student ${userScore.user} using endpoint ${endpoint}. Response: `});
            });
        });
        if (chunkIndex + chunkSize < scores.length) {
            setTimeout(buildRequests, 1000, chunkIndex + chunkSize);
        }
    }
    // Start recursive iteration
    buildRequests(0);
}

defer(function() {
    'use strict';

    // Only add the import button if a rubric is appearing
    if ($('#rubric_summary_holder').length > 0) {
        $("body").append($('<div id="import_rubric_popup_dialog" title="Import Rubric Scores"></div>'));
        $("body").append($('<div id="import_rubric_dialog" title="Import Rubric Scores"></div>'));
        $("body").append($('<div id="import_rubric_progress" title="Import Rubric Scores"><p>Importing rubric scores. Do not navigate from this page.</p><div id="import_rubric_bar"></div></div>'));
        $('#gradebook_header div.statsMetric').append('<button type="button" class="Button" id="import_rubric_btn">Import Rubric Scores</button>');
        $('#import_rubric_btn').click(function() {
            openImportDialog(function(importFile, successCallback) {
                // parse CSV
                Papa.parse(importFile, {
                    header: true,
                    dynamicTyping: false,
                    complete: function(results) {
                        validateData(results.data, successCallback);
                    }
                });
            }, importScores);
        });
    }
});
