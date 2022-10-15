// ==UserScript==
// @name         Comments Importer
// @namespace    https://github.com/UCBoulder
// @description  Bulk import assignment comments into the Canvas gradebook.
// @match        https://*/courses/*/gradebook
// @grant        none
// @require      https://cdnjs.cloudflare.com/ajax/libs/PapaParse/5.1.0/papaparse.min.js
// @run-at       document-idle
// @version      1.1.4
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

defer(function() {
    'use strict';

    // utility function for downloading an error report
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

    // prep jquery info dialog
    $("body").append($('<div id="comments_dialog" title="Import Comments"></div>'));
    $("#comments_dialog").dialog({ autoOpen: false });
    function popUp(text) {
        $("#comments_dialog").html(`<p>${text}</p>`);
        $("#comments_dialog").dialog('open');
    }

    // prep jquery confirm dialog
    $("body").append($('<div id="comments_modal" title="Import Comments"></div>'));
    $("#comments_modal").dialog({ modal: true, autoOpen: false });
    function confirm(text, callback) {
        $("#comments_modal").html(`<p>${text}</p>`);
        $("#comments_modal").dialog({
            buttons: {
                "Confirm": function() {
                    $(this).dialog("close");
                    callback(true);
                },
                "Cancel": function() {
                    $(this).dialog("close");
                    callback(false);
                }
            }
        });
        $("#comments_modal").dialog('open');
    }

    // prep jquery progress dialog
    $("body").append($('<div id="comments_progress" title="Import Comments"><p>Importing comments. Do not navigate from this page.</p><div id="comments_bar"></div></div>'));
    $("#comments_progress").dialog({ buttons: {}, autoOpen: false });
    function showProgress(amount) {
        if (amount === 100) {
            $("#comments_progress").dialog("close");
        } else {
            $("#comments_bar").progressbar({ value: amount });
            $("#comments_progress").dialog("open");
        }
    }

    // add choose file button to gradebook
    var importDiv = $(`<div style="padding-top:10px">
<label for="comments_file">Import comments: </label>
<input type="file" id="comments_file"/>
</div>`);
    $("div.gradebook-menus").append(importDiv);

    // handle when file is selected
    $('#comments_file').change(function(evt) {
        $("#comments_file").hide();
        // parse CSV
        Papa.parse(evt.target.files[0], {
            header: true,
            dynamicTyping: false,
            complete: function(results) {
                $("#comments_file").val('');
                var data = results.data;
                var referral = ' Visit <a href="https://oit.colorado.edu/services/teaching-learning-applications/canvas/enhancements-integrations/enhancements#oit" target="_blank">Canvas - Enhancements</a> for formatting guidelines.';
                if (data.length < 1) {
                    popUp("ERROR: File should contain a header row and at least one data row." + referral);
                    $("#comments_file").show();
                    return;
                }
                if (!Object.keys(data[0]).includes("SIS User ID")) {
                    popUp("ERROR: No 'SIS User ID' column found." + referral);
                    $("#comments_file").show();
                    return;
                }
                if (Object.keys(data[0]).length < 2) {
                    popUp("ERROR: Header row should have a 'SIS User ID' column and at least one assignment column." + referral);
                    $("#comments_file").show();
                    return;
                }

                // build requests
                var requests = [];
                for (const row of data) {
                    var student = row["SIS User ID"];
                    for (const assignment of Object.keys(row)) {
                        if (assignment === "SIS User ID") {
                            continue;
                        }
                        // extract assignment id from assignment header
                        var idWithParens = assignment.match(/\(\d+\)$/);
                        if (!idWithParens) {
                            popUp(`ERROR: "${assignment}" is not a properly formatted assignment name.` + referral);
                            $("#comments_file").show();
                            return;
                        }
                        var assignId = idWithParens[0].match(/\d+/)[0];
                        var comment = row[assignment];
                        if (!comment || !comment.trim()) {
                            continue;
                        }
                        // extract course id from url
                        var courseId = window.location.href.split('/')[4];
                        // build api url
                        var subUrl = `/api/v1/courses/${courseId}/assignments/${assignId}/submissions/${student}`;
                        // build request and canned error message in case it fails
                        requests.push({
                            request: {
                                url: subUrl,
                                type: "PUT",
                                data: {"comment[text_comment]": comment},
                                dataType: "text" },
                            error: `Failed to post comment for student ${student} and assignment ${assignment} using endpoint ${subUrl}. Response: `
                        });
                    }
                }

                // confirm before proceeding
                confirm(
                    `You are about to post ${requests.length} new comments. This cannot be undone. Are you sure you wish to proceed?`,
                    function(confirmed) {
                        if (confirmed) {

                            // send requests in chunks of 10 every second to avoid rate-limiting
                            var errors = [];
                            var completed = 0;
                            var chunkSize = 10;
                            function sendChunk(i) {
                                for (const request of requests.slice(i, i+chunkSize)) {
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
                                    $("#comments_file").show();
                                    if (errors.length > 0) {
                                        popUp(`Import complete. WARNING: ${errors.length} comments failed to import. See errors.txt for details.
`);
                                        saveText(errors, "errors.txt");
                                    } else {
                                        popUp("All comments imported successfully!");
                                    }
                                }
                            }
                            // actually starts the recursion
                            sendChunk(0);
                        } else {
                            // confirmation was dismissed
                            $("#comments_file").show();
                        }
                    });
            }
        });
    });
});
