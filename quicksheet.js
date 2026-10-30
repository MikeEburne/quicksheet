/*
 * Quick Sheet - shared logic used by the ribbon button and the task pane.
 * This is the Office.js equivalent of the VBA in NewSheet.bas.
 */

/* global Excel */

// Excel sheet names: max 31 characters, none of  \ / ? * [ ] :  and not blank.
function cleanSheetName(name) {
  var n = String(name || "").replace(/[\\\/\?\*\[\]:]/g, " ").replace(/^'+|'+$/g, "").trim();
  if (!n) n = "New Sheet";
  return n.slice(0, 31).trim();
}

// "New Sheet" -> "New Sheet (2)" -> "New Sheet (3)" ... until the name is free.
function uniqueSheetName(base, existing) {
  var taken = {};
  existing.forEach(function (s) {
    taken[s.toLowerCase()] = true;
  });
  var name = cleanSheetName(base);
  var i = 2;
  while (taken[name.toLowerCase()]) {
    var suffix = " (" + i + ")";
    name = cleanSheetName(base).slice(0, 31 - suffix.length).trim() + suffix;
    i++;
  }
  return name;
}

// Adds a sheet at the end, writes a header, selects it. Returns the name used.
async function createSheet(requestedName) {
  return Excel.run(async function (context) {
    var sheets = context.workbook.worksheets;
    sheets.load("items/name");
    await context.sync();

    var name = uniqueSheetName(
      requestedName || "New Sheet",
      sheets.items.map(function (s) {
        return s.name;
      })
    );

    var ws = sheets.add(name);
    var header = ws.getRange("A1");
    header.values = [["Created by Quick Sheet on " + new Date().toLocaleString()]];
    header.format.font.bold = true;
    ws.activate();
    ws.getRange("A2").select();
    await context.sync();
    return name;
  });
}

// let Node load this file for the tests
if (typeof module !== "undefined") module.exports = { cleanSheetName: cleanSheetName, uniqueSheetName: uniqueSheetName };
