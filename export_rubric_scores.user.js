// ==UserScript==
// @name         Export Rubric Scores
// @namespace    https://github.com/UCBoulder
// @description  Export all rubric criteria scores for an assignment to a CSV
// @match        https://*/courses/*/gradebook/speed_grader?*
// @grant        none
// @require      https://code.jquery.com/jquery-3.6.0.min.js
// @require      https://code.jquery.com/ui/1.14.1/jquery-ui.min.js
// @run-at       document-idle
// @version      1.2.6
// ==/UserScript==

/* globals $ */

function defer(method) {
    if (typeof $ !== 'undefined') {
        method();
    } else {
        setTimeout(function () { defer(method); }, 100);
    }
}

function waitForElement(selector, callback) {
    if ($(selector).length) {
        callback();
    } else {
        setTimeout(function () {
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

function getRemainingPages(nextUrl, listSoFar, callback) {
    $.getJSON(nextUrl, function (responseList, textStatus, jqXHR) {
        let nextLink = null;
        const linkHeader = jqXHR.getResponseHeader("link");
        if (linkHeader) {
            $.each(linkHeader.split(','), function (linkIndex, linkEntry) {
                if (linkEntry.split(';')[1].includes('rel="next"')) {
                    nextLink = linkEntry.split(';')[0].slice(1, -1);
                }
            });
        }
        if (!nextLink) {
            callback(listSoFar.concat(responseList));
        } else {
            getRemainingPages(nextLink, listSoFar.concat(responseList), callback);
        }
    }).fail(function (jqXHR, textStatus, errorThrown) {
        popUp(`ERROR ${jqXHR.status} while retrieving data from Canvas. Url: ${nextUrl}<br/><br/>Please refresh and try again.`);
        window.removeEventListener("error", showError);
    });
}

function csvEncode(string) {
    if (string && (string.includes('"') || string.includes(','))) {
        return '"' + string.replace(/"/g, '""') + '"';
    }
    return string;
}

function showError(event) {
    popUp("JavaScript error: " + event.message);
    window.removeEventListener("error", showError);
}

defer(function () {
    'use strict';

    var saveText = (function () {
        var a = document.createElement("a");
        document.body.appendChild(a);
        a.style = "display: none";
        return function (textArray, fileName) {
            var blob = new Blob(textArray, { type: "text" }),
                url = window.URL.createObjectURL(blob);
            a.href = url;
            a.download = fileName;
            a.click();
            window.URL.revokeObjectURL(url);
        };
    })();

    $("body").append($('<div id="export_rubric_dialog" title="Export Rubric Scores"></div>'));

    try {
        if ($('#rubric_summary_holder').length > 0) {
            $('#gradebook_header div.statsMetric').append('<button type="button" class="Button" id="export_rubric_btn">Export Rubric Scores</button>');
        } else {
            console.warn("Rubric summary holder not found. Export button not inserted.");
        }
    } catch (e) {
        popUp("DOM error while trying to insert export button: " + e.message);
    }

    $('#export_rubric_btn').click(function () {
        try {
            popUp("Exporting scores, please wait...");
            window.addEventListener("error", showError);

            const courseId = window.location.href.split('/')[4];
            const urlParams = window.location.href.split('?')[1].split('&');
            const assignId = urlParams.find(i => i.split('=')[0] === "assignment_id").split('=')[1];

            $.getJSON(`/api/v1/courses/${courseId}/assignments/${assignId}`, function (assignment) {
                getAllPages(`/api/v1/courses/${courseId}/enrollments?per_page=100`, function (enrollments) {
                    getAllPages(`/api/v1/courses/${courseId}/assignments/${assignId}/submissions?include[]=rubric_assessment&per_page=100`, function (submissions) {

                        if (!('rubric_settings' in assignment)) {
                            popUp(`ERROR: No rubric settings found at /api/v1/courses/${courseId}/assignments/${assignId}.<br/><br/>
                                This is likely due to a Canvas bug where a rubric has entered a "soft-deleted" state.
                                Please use the <a href="https://community.canvaslms.com/t5/Canvas-Admin-Blog/Undeleting-things-in-Canvas/ba-p/267116">Undelete feature</a>
                                to restore the rubric associated with this assignment or contact Canvas Support.`);
                            return;
                        }

                        const hidePoints = assignment.rubric_settings.hide_points;
                        const hideRatings = assignment.rubric_settings.free_form_criterion_comments;

                        if (hidePoints && hideRatings) {
                            popUp("ERROR: This rubric is configured to use free-form comments instead of ratings AND to hide points, so there is nothing to export!");
                            return;
                        }

                        let critOrder = {};
                        let critRatingDescs = {};
                        let header = "Student Name,Student ID,Posted Score,Attempt Number";
                        $.each(assignment.rubric, function (critIndex, criterion) {
                            critOrder[criterion.id] = critIndex;
                            if (!hideRatings) {
                                critRatingDescs[criterion.id] = {};
                                $.each(criterion.ratings, function (i, rating) {
                                    critRatingDescs[criterion.id][rating.id] = rating.description;
                                });
                                header += ',' + csvEncode('Rating: ' + criterion.description);
                            }
                            if (!hidePoints) {
                                header += ',' + csvEncode('Points: ' + criterion.description);
                            }
                        });
                        header += '\n';

                        var csvRows = [header];
                        $.each(submissions, function (subIndex, submission) {
                            const enrollment = enrollments.find(i => i.user_id === submission.user_id);
                            if (enrollment && enrollment.user) {
                                const user = enrollment.user;
                                let row = `${user.name},${user.sis_user_id},${submission.score},${submission.attempt}`;

                                let crits = [];
                                let critIds = [];
                                if (submission.rubric_assessment != null) {
                                    $.each(submission.rubric_assessment, function (critKey, critValue) {
                                        crits.push({
                                            id: critKey,
                                            points: critValue.points ?? null,
                                            rating: hideRatings ? null : (critRatingDescs[critKey]?.[critValue.rating_id] ?? null)
                                        });
                                        critIds.push(critKey);
                                    });
                                }

                                $.each(critOrder, function (critKey) {
                                    if (!critIds.includes(critKey)) {
                                        crits.push({ id: critKey, points: null, rating: null });
                                    }
                                });

                                crits.sort(function (a, b) { return critOrder[a.id] - critOrder[b.id]; });
                                $.each(crits, function (critIndex, criterion) {
                                    if (!hideRatings) {
                                        row += `,${csvEncode(criterion.rating)}`;
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
                        const fileName = `Rubric_Scores_${assignment.name.replace(/[^a-zA-Z0-9]/g, '_')}.csv`;
                        saveText(csvRows, fileName);
                        window.removeEventListener("error", showError);
                    });
                });
            }).fail(function (jqXHR, textStatus, errorThrown) {
                popUp(`ERROR ${jqXHR.status} while retrieving assignment data from Canvas. Please refresh and try again.`);
                window.removeEventListener("error", showError);
            });

        } catch (e) {
            popUp("Unexpected error: " + e.message);
            window.removeEventListener("error", showError);
        }
    });
});
