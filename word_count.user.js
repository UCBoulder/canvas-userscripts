// ==UserScript==
// @name         Speedgrader Word Count
// @namespace    https://github.com/UCBoulder
// @description  Displays a word count for documents within Speedgrader
// @match        https://canvadocs.instructure.com/*
// @require      https://code.jquery.com/jquery-1.12.4.min.js
// @grant        none
// @run-at       document-idle
// @version      1.0.4
// ==/UserScript==

/* globals $ */

function getWordCount() {
    var wc = 0;
    $('div.textLayer > span').each(function(divIndex) {
        wc += $(this).text().split(/\s+\b/).length;
    });
    return wc;
}

function updateWordCount(idleFrames, lastWC) {
    var wc = getWordCount();
    if (wc === lastWC) {
        idleFrames++;
    } else {
        $('#swc_display').text(`Word count: ${wc}`);
        idleFrames = 0;
    }
    // If WC doesn't update for 5 seconds, stop checking
    if (idleFrames < 10) {
        setTimeout(function () { updateWordCount(idleFrames, wc); }, 500);
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

(function() {
    'use strict';

    waitForElement('div.textLayer', function() {
        $('#App > nav > div > div.ViewerControls--title').append('<span id="swc_display" style="margin-right: 1em">Word count:</span>');
        updateWordCount(0, 0);
    });
})();
