// ==UserScript==
// @name         Copy & Offset Assignment Dates
// @namespace    https://github.com/CUBoulder-OIT
// @description  Copy all "Assign to" dates from one assignment to another, with offset.
// @include      https://canvas.*.edu/courses/*/assignments/*
// @include      https://*.*instructure.com/courses/*/assignments/*
// @require      https://cdnjs.cloudflare.com/ajax/libs/moment.js/2.22.2/moment-with-locales.min.js
// @require      https://cdnjs.cloudflare.com/ajax/libs/moment-timezone/0.5.23/moment-timezone-with-data.min.js
// @grant        none
// @run-at       document-idle
// @version      1.0.2
// ==/UserScript==

/* globals $ moment */

// wait until the window jQuery is loaded
function defer(method) {
    if (typeof $ !== 'undefined') {
        method();
    }
    else {
        setTimeout(function() { defer(method); }, 100);
    }
}

defer(function() {
    'use strict';

    // prep jquery info dialog
    $("body").append($('<div id="caod_dialog" title="Copy & Offset Dates"></div>'));
    function popUp(text) {
        $("#caod_dialog").html(`<p>${text}</p>`);
        $("#caod_dialog").dialog();
    }
    function popClose() {
        $("#caod_dialog").dialog("close");
    }

    // prep jquery form dialog
    $("body").append($('<div id="caod_form_dialog" title="Copy & Offset Dates">'));
    function openForm(assignments, srcAssign, infoMsg, selectedUnits, submitAction) {
        // parse arguments that specify options
        var infoHtml = "";
        if (infoMsg) {
            infoHtml = `<p><br>${infoMsg}</p>`;
        }
        var unitsOptions = `<option value="weeks">weeks</option>
<option value="days">days</option>`;
        if (selectedUnits === "weeks") {
            unitsOptions = `<option value="weeks" selected>weeks</option>
<option value="days">days</option>`
        } else if (selectedUnits === "days") {
            unitsOptions = `<option value="weeks">weeks</option>
<option value="days" selected>days</option>`
        }

        // built select options from assignment list
        var assignOptions = '<option value="none_chosen"></option>';
        $.each(assignments, function(index, assignment) {
            if (assignment.id !== srcAssign.id) {
                assignOptions = `${assignOptions}
<option value="${assignment.id}">${assignment.name}</option>`;
            }
        });

        // build HTML
        $("#caod_form_dialog").html(`
<p>
Copy all "Assign to" options and dates from the current assignment to another assignment. This will overwrite ALL "Assign to" options and dates for the destination assignment.
</p>
<hr>
<p>Copy from: ${srcAssign.name}</p>
<form id="caod_form">
<label>Copy to</label>
<select name="caod_assignment">
${assignOptions}
</select>
<hr>
<label>Offset Type</label>
<select name="caod_offset_type">
${unitsOptions}
</select>
<br>
<label>Offset Amount</label>
<input type="number" autocomplete="off" name="caod_offset">
<br>
<br>
<input type="submit" value="Submit" class="btn btn-primary">
${infoHtml}
</form>`);

        // on submit, send form data to submitAction
        $("#caod_form").submit(function(event) {
            event.preventDefault();
            $("#caod_form_dialog").dialog("close");
            submitAction($("#caod_form").serialize());
        });
        $("#caod_form_dialog").dialog({width: "375px"});
    }

    // utility function for getting value based on field name from serialized form data
    function getFormValue(name, serial) {
        for (const elem of serial.split('&')) {
            var pair = elem.split('=');
            if (pair[0] === name) {
                return pair[1];
            }
        }
    }

    // respond to submitted form and actually perform the copy action
    function processForm(formData, srcAssign, allAssigns, courseData) {
        popUp("Copying dates. Please do not navigate away from this page.");

        // extract data from submitted form
        var destAssign = getFormValue('caod_assignment', formData);
        var offset = getFormValue('caod_offset', formData) || 0;
        var units = getFormValue('caod_offset_type', formData);
        if (destAssign === "none_chosen") {
            popClose();
            openForm(allAssigns, srcAssign, "ERROR: You must choose a destination assignment", units, function(formData) { processForm(formData, srcAssign, allAssigns, courseData); });
            return;
        }

        // copy and offset overrides for the parameters
        var overrides = [];
        $.each(srcAssign.overrides, function (srcIndex, srcElem) {
            var override = JSON.parse(JSON.stringify(srcElem));
            delete override.id;
            delete override.assignment_id;
            for (const dateType of ["due_at", "unlock_at", "lock_at"]) {
                if (override[dateType]) {
                    var dt = moment.tz(override[dateType], courseData.time_zone);
                    dt.add(offset, units);
                    override[dateType] = dt.utc().format();
                } else {
                    delete override[dateType];
                }
            }
            overrides.push(override);
        });
        var params = { "assignment[assignment_overrides]": overrides };

        // get "base" assignment dates/settings
        for (const dateType of ["due_at", "unlock_at", "lock_at"]) {
            if (srcAssign[dateType]) {
                var dt = moment.tz(srcAssign[dateType], courseData.time_zone);
                dt.add(offset, units);
                params[`assignment[${dateType}]`] = dt.utc().format();
            } else {
                params[`assignment[${dateType}]`] = "";
            }
        }
        params["assignment[only_visible_to_overrides]"] = srcAssign.only_visible_to_overrides;


        // send the request, reload form on success
        $.ajax({
            url: `/api/v1/courses/${courseData.id}/assignments/${destAssign}`,
            type: "PUT",
            data: params,
            dataType: "text"
        }).fail(function (jqXHR, textStatus, errorThrown) {
            popUp(`ERROR ${jqXHR.status} while updating assignment. Please refresh and try again.`);
        }).done(function () {
            popClose();
            openForm(allAssigns, srcAssign, "Success! Dates copied.", units, function(formData) { processForm(formData, srcAssign, allAssigns, courseData); });
        });
    }

    // add offset dates button to assignment page
    if ($("#assignment_show > div.content-box").length) {
         // regular assignment
         $("#assignment_show > div.content-box").prepend($('<button type="button" id="caod_button" class="btn Button" role="button">Copy & Offset Dates</button>'));
    } else {
        // external tool assignment
        $("#content > div.tool_content_wrapper").after($('<div class="content-box"><button type="button" id="caod_button" class="btn Button" role="button">Copy & Offset Dates</button></div>'));
    }
    $("#caod_button").click(function() {
        popUp("Loading assignments. Please wait.");

        // get data necessary for rendering the form
        var courseId = window.location.href.split('/')[4];
        var srcAssignId = window.location.href.split('/')[6];
        $.getJSON(`/api/v1/courses/${courseId}/assignments?per_page=100`, function(allAssigns) {
            $.getJSON(`/api/v1/courses/${courseId}`, function(courseData) {
                $.getJSON(`/api/v1/courses/${courseData.id}/assignments/${srcAssignId}?include[]=overrides&override_assignment_dates=false`, function(srcAssign) {
                    popClose();
                    openForm(allAssigns, srcAssign, null, null, function(formData) { processForm(formData, srcAssign, allAssigns, courseData); });
                });
            });
        });
    });
});
