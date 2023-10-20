/**
 * --------------------------------------------------------------------------------------
 * DMH 10/17/2023
 *  Provide extra functionality for the Fact Tracker Worksheet, like Help and Reference Field Highlighting
 *  This is a container script bound to the Fact Tracker Worksheet Pattern
 *    https://docs.google.com/spreadsheets/d/1INHBcR1jwrG8DgkfIwDFE0HWYU0kUUBmfJ6wLYupEBA/edit#gid=1562645589
 *  Documentation
 *    Spreadsheet Service:                https://developers.google.com/apps-script/reference/spreadsheet
 *    Creating Menu Items:                https://developers.google.com/apps-script/guides/menus
 *    Class UI:                           https://developers.google.com/apps-script/reference/base/ui.html
 *    Banding Class:                      https://developers.google.com/apps-script/reference/spreadsheet/banding
 *    spreadsheet.getBandings:            https://developers.google.com/apps-script/reference/spreadsheet/spreadsheet#getbandings
 *    range.applyRowBandings:             https://developers.google.com/apps-script/reference/spreadsheet/range#applyrowbanding
 * 
 * --------------------------------------------------------------------------------------
 */

//--------------------------------------------------------------------------------------------------
function showFactTrackerGuide()
{
  GrandJuryLibrary.message('showFactTrackerGuide() started...')

  GrandJuryLibrary.showFactTrackerGuide()
  GrandJuryLibrary.message('showFactTrackerGuide() ended')
}


//--------------------------------------------------------------------------------------------------
function showRowFieldHighlights()
{
  GrandJuryLibrary.message('showRowFieldHighlights() started...')
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const row         = spreadsheet.getActiveCell().getRow()
  GrandJuryLibrary.showRowFieldHighlights(spreadsheet, row)
  GrandJuryLibrary.message('showRowFieldHighlights() started...')
}


//--------------------------------------------------------------------------------------------------
function clearAllHighlights()
{
  GrandJuryLibrary.message('clearAllHighlights() started...')
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  GrandJuryLibrary.clearAllHighlights(spreadsheet)
  GrandJuryLibrary.message('clearAllHighlights() ended')
}

//--------------------------------------------------------------------------------------------------
function hideRowFieldHighlights()
{
  GrandJuryLibrary.message('hideRowFieldHighlights() started...')
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet()
  const row         = spreadsheet.getActiveCell().getRow()
  GrandJuryLibrary.hideRowFieldHighlights(spreadsheet, row)
  GrandJuryLibrary.message('hideRowFieldHighlights() ended')
}


//--------------------------------------------------------------------------------------------------
function createMenuItems()
{
  const ui  = SpreadsheetApp.getUi()
  ui.createMenu('Fact Tracker Extras')
    .addItem('Show Fact Tracker Guide',   'showFactTrackerGuide')
    .addItem('Show Row Field Highlights', 'showRowFieldHighlights')
    .addItem('Hide Row Field Highlights', 'hideRowFieldHighlights')
    .addItem('Clear All Highlights',      'clearAllHighlights')
    .addToUi()

}

//--------------------------------------------------------------------------------------------------
function onOpen(e)
{
  GrandJuryLibrary.message('onOpen() started...')
  createMenuItems()
  GrandJuryLibrary.message('onOpen() ended')
}

//--------------------------------------------------------------------------------------------------
function messageStartEnd()
{
  //  We can use these to measure function duration
  GrandJuryLibrary.message('() started...')
  GrandJuryLibrary.message('() ended')
}
