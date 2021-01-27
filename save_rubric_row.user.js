// ==UserScript==
// @name         Save Rubric Row
// @namespace    https://github.com/CUBoulder-OIT
// @description  When in Speedgrader, save comments/score for just one row of a rubric at a time.
// @include      https://canvas.*.edu/courses/*/gradebook/speed_grader?*
// @include      https://*.*instructure.com/courses/*/gradebook/speed_grader?*
// @grant        none
// @run-at       document-idle
// @version      1.0.5
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
    $("#srr_dialog").html(`<p>${text}</p>`);
    $("#srr_dialog").dialog();
}

function saveCriterion(rowIndex, callback) {
    // Get some initial data from the current URL
    var courseId = window.location.href.split('/')[4];
    var urlParams = window.location.href.split('?')[1].split('&');
    var assignId, studentId;
    $.each(urlParams, function(i, param) {
        switch(param.split('=')[0]) {
            case "assignment_id":
                assignId = param.split('=')[1];
                break;
            case "student_id":
                studentId = param.split('=')[1];
                break;
        }
    });

    // Get rubrics ids from the assignments API
    $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`, function(assignment) {

        // And get existing rubric assessment from the submissions API
        $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}/submissions/${studentId}?include[]=rubric_assessment`, function(submission) {

            // Pre-load params with existing rubric assessment data
            var params = {};
            $.each(submission.rubric_assessment, function(rowKey, rowValue) {
                $.each(rowValue, function(cellKey, cellValue) {
                    params[`rubric_assessment[${rowKey}][${cellKey}]`] = cellValue;
                });
            });

            // Get Canvas's identifier for the row to be updated
            var rowId = assignment.rubric[rowIndex].id;

            // Determine the index of the selected tier
            var tier;
            $($('tr[data-testid="rubric-criterion"]')[rowIndex]).find('.rating-tier').each(function(tierIndex) {
                if ($(this).hasClass("selected")) {
                    tier = tierIndex;
                }
            });
            // Set the new rating to match the selected tier
            if (tier === undefined) {
                // Make sure the rating is cleared if none is selected
                params[`rubric_assessment[${rowId}][rating_id]`] = undefined;
            } else {
                params[`rubric_assessment[${rowId}][rating_id]`] = assignment.rubric[rowIndex].ratings[tier].id;
            }

            // If points are hidden, we will need to set them based on the chosen rating
            if (assignment.rubric_settings.hide_points) {
                params[`rubric_assessment[${rowId}][points]`] = assignment.rubric[rowIndex].ratings[tier].points;
            } else {
                // Otherwise, set points based on what's entered
                const score = $($('td[data-testid="criterion-points"] input')[rowIndex]).val();
                if (isNaN(score)) {
                    // Make sure the score is cleared if blank or invalid
                    params[`rubric_assessment[${rowId}][points]`] = undefined;
                    // Clear the field as well for the sake of clarity
                    $($('td[data-testid="criterion-points"] input')[rowIndex]).val('');
                } else {
                    params[`rubric_assessment[${rowId}][points]`] = score;
                }
            }

            // Get entered comments (comments should never be undefined)
            var comments = $($('#rubric_full tr[data-testid="rubric-criterion"]')[rowIndex]).find('textarea').val();
            if (comments === undefined) {
                comments = "";
            }
            params[`rubric_assessment[${rowId}][comments]`] = comments;

            // Send the updated rubric assessment
            $.ajax({
                url: `/api/v1/courses/${courseId}/assignments/${assignId}/submissions/${studentId}`,
                type: 'PUT',
                data: params,
                dataType: "text"
            }).fail(function (jqXHR, textStatus, errorThrown) {
                popUp(`ERROR ${jqXHR.status} while saving score. Please refresh and try again.`);
                callback(false);
            }).done(function () {
                callback(true);
            });
        });
    });
}

defer(function() {
    'use strict';

    // prep jquery info dialog
    $("body").append($('<div id="srr_dialog" title="Save Rubric Row"></div>'));

    waitForElement('#rubric_assessments_list_and_edit_button_holder > div > button', function() {
        $('#rubric_assessments_list_and_edit_button_holder > div > button').click(function() {
            // Add in buttons if they don't already exist
            if ($('#save_row_0').length === 0) {
                $('td[data-testid="criterion-points"]').each(function(index) {
                    var saveBtn = $(`<button type="button" class="Button Button--primary" id="save_row_${index}" style="margin-top: 0.375em">Save Row</button>`)
                    saveBtn.click(function() {
                        saveCriterion(index, function(success) {
                            if (success) {
                                $($('td[data-testid="criterion-points"]')[index]).append(`<span style="display: block; margin-top: 0.375em" id=save_row_${index}_alert role="alert">Saved!</span>`);
                                setTimeout(function() {
                                    $(`#save_row_${index}_alert`).remove();
                                }, 1500);
                            }
                        });
                    });
                    $(this).append(saveBtn);
                });
            }
        });
    });
});
