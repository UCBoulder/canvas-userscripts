// ==UserScript==
// @name         Custom Column Manager
// @namespace    https://github.com/UCBoulder
// @description  Add buttons to create and delete custom gradebook columns in Canvas LMS.
// @include      https://canvas.*.edu/*/gradebook
// @include      https://*.*instructure.com/*/gradebook
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

function rmColumn() {
    var colName = prompt("Enter the name of the column you wish to delete.\n\nWARNING: All data in this column will be permanently lost!", "");
    if (colName) {
        //extract course id from url
        var courseId = window.location.href.split('/')[4];
        //build the url to the custom columns api endpoint
        var colUrl = "/api/v1/courses/" + courseId + "/custom_gradebook_columns";
        //get a list of existing custom columns
        $.ajax({url: colUrl, dataType: "text"}).done(function(data) {
            var columns = JSON.parse(data.substr(data.indexOf('['),data.length));
            for (var i = 0; i < columns.length; i++) {
                //Match the column that should be deleted by name
                if (columns[i].title === colName) {
                    //delete the column
                    $.ajax({
                        url: colUrl + "/" + columns[i].id,
                        type: "DELETE",
                        dataType: "text"
                    }).done(function(data) {
                        alert('"' + columns[i].title + '" deleted!');
                        location.reload();
                    }).fail(function(jqXHR, textStatus, errorThrown) {
                        alert("Error when deleting column: " + errorThrown);
                    });
                    return;
                }
            }
            //for loop completed without finding the column
            alert('Could not find a column named "' + colName + '".');
        }).fail(function(jqXHR, textStatus, errorThrown) {
            alert("Error when checking column names: " + errorThrown);
        });
    }
}

function addColumn() {
    var colName = prompt("Enter the name of the new column you wish to create.", "");
    if (colName) {
        //extract course id from url
        var courseId = window.location.href.split('/')[4];
        //build the url to the custom columns api endpoint
        var colUrl = "/api/v1/courses/" + courseId + "/custom_gradebook_columns";
        //get a list of existing columns
        $.get(colUrl, function(data) {
            for (var col in data) {
                //if a column with this name already exists, quit
                if (col.title === colName) {
                    alert('A column named "' + colName + '" already exists.');
                    return;
                }
            }
        });
        //create the column
        $.ajax({
            url: colUrl,
            type: "POST",
            data: {"column[title]": colName},
            dataType: "text"
        }).done(function(data) {
            alert('"' + colName + '" created!');
            location.reload();
        }).fail(function(jqXHR, textStatus, errorThrown) {
            alert("Error when creating column: " + errorThrown);
        });
    }
}

defer(function() {
    'use strict';
    var colDiv = $(`<div>
<button class="btn" role="button" aria-label="Add Custom Column" id="addColumnBtn"><i class="icon-plus"/> Column</button>
<button class="btn" role="button" aria-label="Delete Custom Column" id="rmColumnBtn">Delete Column</button>
</div>`);
    $("#gradebook-actions").prepend(colDiv);
    $("#rmColumnBtn").click(rmColumn);
    $("#addColumnBtn").click(addColumn);
});
