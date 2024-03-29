﻿/*
* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license at the bottom of this file.
*/
// This function is run when the app is ready to start interacting with the host application
// It ensures the DOM is ready before adding click handlers to buttons
Office.initialize = function (reason) {
    $(document).ready(function () {
        var panel;
        $('.tabs a').click(function () {
            // save $(this) in a variable for efficiency
            var $this = $(this);

            // hide panels
            $('.panel').hide();

            // remove the active state from the tabs if already set
            $('.tabs a.active').removeClass('active');

            // add active state to the tab
            $this.addClass('active').blur();
            // retrieve href from link (the id of panel to display)
            panel = $this.attr('href');
            // show panel
            $(panel).fadeIn(250);
            return (false);
        }); // end .tabs

        // open first tab
        $('.tabs li:first a').click();


        $('#setDataBtn').click(function () { setData($(panel).children('p')); });
            return (false);

    });
};

// Writes data from textbox to the current selection in the document
function setData(elementId) {
    Office.context.document.setSelectedDataAsync($(elementId).text());
}

// *********************************************************
//
// Excel-Add-in-Create-tab-Menu, https://github.com/OfficeDev/Excel-Add-in-Create-tab-Menu
//
// Copyright (c) Microsoft Corporation
// All rights reserved.
//
// MIT License:
// Permission is hereby granted, free of charge, to any person obtaining
// a copy of this software and associated documentation files (the
// "Software"), to deal in the Software without restriction, including
// without limitation the rights to use, copy, modify, merge, publish,
// distribute, sublicense, and/or sell copies of the Software, and to
// permit persons to whom the Software is furnished to do so, subject to
// the following conditions:
//
// The above copyright notice and this permission notice shall be
// included in all copies or substantial portions of the Software.
//
// THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND,
// EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF
// MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND
// NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE
// LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION
// OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION
// WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
//
// *********************************************************