/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global console, document, Excel, Office */

const exclusionListUrl = "https://raw.githubusercontent.com/dmcwee/dmcwee.github.io/master/MDEExclusionChecker/exclusionList/mde-exclusions.json";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = run;
    document.getElementById("clearLog").onclick = clearLog;
  }
});

export async function run() {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      log("Staring Analysis");

      var range = context.workbook.getSelectedRange();
      range.load("address");
      range.load("values");

      log("Before Await");
      await context.sync();
      log(`range values: ${JSON.stringify(range.values)}`);

      //get the mde exclusions
      log(`Calling ${exclusionListUrl}`);
      let response = await fetch(
        exclusionListUrl
      );
      if (response.ok) {
        log(`Response was OK.`);
        var json = await response.json();
        log(`Response: ${JSON.stringify(json)}`);

        var newRange = validateRange(range, json);

        log(`newRange: ${JSON.stringify(newRange)}`);

        range.values = newRange;
        range.format.autofitColumns();

        context.sync();

        log(`DONE!`);
        return;
      } else {
        log("There was a problem loading the list of Exclusion Paths");
      }
    });
  } catch (error) {
    console.error(error);
  }
}

export function clearLog() {
  document.getElementById("logMessages").value = "";
}

function log(message) {
  console.log(`Log: ${message}`);
  document.getElementById("logMessages").value += `${message}\n`;
}

function validateRange(range, defaultList) {
  var validatedRange = [];
  for (let i = 0; i < range.values.length; ++i) {
    let exclusion = range.values[i];
    validatedRange.push(compareExclusionToDefaultList(exclusion[0], defaultList));
    log(`validated range after push is now ${JSON.stringify(validatedRange)}`);
  }
  return validatedRange;
}

function compareExclusionToDefaultList(exclusion, defaultList) {
  for (let i = 0; i < defaultList.length; ++i) {
    let defaultRule = defaultList[i];
    log(`Comparing ${JSON.stringify(exclusion)} agaist ${JSON.stringify(defaultRule.name)}:${JSON.stringify(defaultRule.paths)}`);
    if (compareExclusionToDefaultPaths(exclusion, defaultRule.paths)) {
      log(` + Match.  Returning ${JSON.stringify([exclusion, JSON.stringify(defaultRule.paths), defaultRule.name])}`);
      return [exclusion, defaultRule.name];
    } else {
      log(` - Match`);
    }
  }
  return [exclusion, "No Matches"];
}

function compareExclusionToDefaultPaths(exclusion, defaultPaths) {
  for (let i = 0; i < defaultPaths.length; ++i) {
    let defaultPath = defaultPaths[i];
    log(`Comparing ${JSON.stringify(exclusion)} against path ${JSON.stringify(defaultPath)}`);
    let result = checkMatch(exclusion, defaultPath);
    if (result) {
      log(`compareExclusionToDefaultPath returning true after match with ${defaultPath}`);
      return true;
    }
  }
  return false;
}

function checkMatch(exclusion, defaultPathRegex) {
  var regex = new RegExp(defaultPathRegex, "ig");
  let result = regex.test(exclusion);
  let match = exclusion.match(regex);

  //log(`RegEx.toString(): ${regex.toString()} - ${exclusion.match(regex)} or ${regex.test(exclusion)}`);
  
  log(`checkMatch returning ${result} and match ${match}`);
  //return regex.test(exclusion);
  return result;
}
