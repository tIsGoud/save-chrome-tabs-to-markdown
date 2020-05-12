/************************

Export Google Chrome tabs to a Markdown file
Version 1.0
April 26, 2020

Usage:
    Run the script when Google Chrome is open.

Customization:
    By default the new file is saved in "Documents".
    Change the variable 'directory' to change the location.
    Example `pathTo` variables: `desktop`, `documents folder`, `home folder`

Requirements:
    - Google Chrome installed
    - Destination directory should exist

Changelog:
    1.00 by @tisgoud, first public version

************************/

// Declare the global variables
var title = ``
var text = ``
var directory = ``
var file = ``

// Get the Computer Name via the Standard Additions of the current app
currentApp = Application.currentApplication()
currentApp.includeStandardAdditions = true

// Get the variable values
title = shortDate(currentApp.currentDate(), true)+ `-chrome-tabs-` + currentApp.systemInfo().computerName
header = `Chrome tabs on ` + currentApp.systemInfo().computerName
directory = currentApp.pathTo(`documents folder`).toString()
file = `${directory}/${title}.md`

getBookmarks(header)
if (writeTextToFile(text, file, true)) {
    currentApp.displayNotification(`All tabs have been saved to ${file}.`, {
        withTitle: `ChromeTabsToMarkdown`})
}

function getBookmarks(headerLine) {
    var window = {}, tab = {}
    var numberOfWindows = 0, numberOfTabs = 0
    var totalTabs = 0
    var i,j

    browser = Application('Google Chrome')

    if (browser.running()) {
        for (i = 0, numberOfWindows = browser.windows.length; i < numberOfWindows; i++) {
            window = browser.windows[i]

            for (j = 0, numberOfTabs = window.tabs.length; j < numberOfTabs; j++) {
                tab = window.tabs[j]

                // Convert title and URL to markdown, empty title is replaced by the URL
                text += '\n- [' + (tab.title().length != 0 ? cleanString(tab.title()) : tab.url()) + ']' + '(' + tab.url() + ')'
            }
            totalTabs += numberOfTabs
            // Add a line between the different windows
            text += '\n\n---\n'
        }
        // Add a title and sub-titles with the date and number of windows and tabs.
        text = '# ' + headerLine + '\n\n' + shortDate(currentApp.currentDate()) + ', ' + totalTabs + ' tabs in ' + numberOfWindows + ' windows\n' +  text
    }
    else {
        text = `no browser running`
    }
}

function shortDate(date, yearFirst = false) {
    var day = ``
    var month = ``
    var year = ``

    day = date.getDate()
    if (day < 10) {
        day = `0` + day;
    }
    month = date.getMonth() + 1
    if (month < 10) {
        month = `0` + month;
    }
    year = date.getFullYear()

    if (yearFirst) {
        return year + `-` + month + `-` + day
    }
    else {
        return day + `-` + month + `-` + year
    }
}

function cleanString(input) {
    var output = ``;

    for (var i = 0; i < input.length; i++) {
      if (input.charCodeAt(i) <= 127) {
        output += input.charAt(i);
      } else {
        output += `&#` + input.charCodeAt(i) + `;`;
      }
    }
    return output;
}

function writeTextToFile(text, file, overwriteExistingContent) {
    try {

        // Convert the file to a string
        var fileString = file.toString()

        // Open the file for writing
        var openedFile = currentApp.openForAccess(Path(fileString), { writePermission: true })

        // Clear the file if content should be overwritten
        if (overwriteExistingContent) {
            currentApp.setEof(openedFile, { to: 0 })
        }

        // Write the new content to the file
        currentApp.write(text, { to: openedFile, startingAt: currentApp.getEof(openedFile) })

        // Close the file
        currentApp.closeAccess(openedFile)

        // Send success notification
        currentApp.displayNotification(`All tabs have been saved to ${file}.`, {
        withTitle: `ChromeTabsToMarkdown`})

        // Return a boolean indicating that writing was successful
        return true
    }
    catch(error) {

        try {
            // Close the file
            currentApp.closeAccess(file)
        }
        catch(error) {
            // Report the error is closing failed
            console.log(`Couldn't close file: ${error}`)

            // Display an alert to the user with a probable cause
            currentApp.displayAlert(`Error in ChromeTabsToMarkdown!`, {
            message: `Please check for non-existing directories: \n${file}.`,
            as: `critical`,
            buttons: [`Oops`],
            givingUpAfter: 10
            })
        }

        // Return a boolean indicating that writing was successful
        return false
    }
}
