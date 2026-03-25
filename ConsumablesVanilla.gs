function populateCombatBuffs() {
  var codeVersion = '2.0.0';
  var confSpreadSheet = SpreadsheetApp.openById('1Xvl3pL_wCbo6LLHUtDx3H9UQbqIgjrfuyBh9RtRyq7w');
  var confSpreadSheetWOTLK = SpreadsheetApp.openById('1QGt6Vv4JBCZ86KQXmBBmCmKTKJbpD4IPX-FH50vXg00');
  var currentVersion = confSpreadSheet.getSheetByName("currentVersion").getRange(1, 1).getValue();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = SpreadsheetApp.getActiveSheet();
  var instructionsSheet = ss.getSheetByName("Instructions");

  try { var lang = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^1.$").useRegularExpression(true).findNext(), 4).getValue(); } catch { }
  var langSheet = confSpreadSheetWOTLK.getSheetByName("langTexts");
  if (lang != null && lang == "English") {
    lang = "EN";
    offset = 1;
  } else if (lang != null && lang == "Deutsch") {
    lang = "DE";
    offset = 2;
  } else if (lang != null && lang == "简体中文") {
    lang = "CN";
    offset = 3;
  } else if (lang != null && lang == "русский") {
    lang = "RU";
    offset = 4;
  } else if (lang != null && lang == "français") {
    lang = "FR";
    offset = 5;
  } else {
    lang = "EN";
    offset = 1;
  }
  var langKeys = langSheet.getRange(1, 1, 1000, 1).getValues().reduce(function (ar, e) { ar.push(e[0]); return ar; }, []);
  var langTrans = langSheet.getRange(1, 1 + offset, 1000, 1).getValues().reduce(function (ar, e) { ar.push(e[0]); return ar; }, []);

  if (currentVersion.indexOf(codeVersion) < 0) {
    SpreadsheetApp.getUi().alert(getStringForLang("sheetOutdated", langKeys, langTrans, "", "", "", ""));
  }

  var firstPlayerNameRow = 9;
  var firstPlayerNameColumn = 2;

  var sheets = ss.getSheets();
  for (var c = sheets.length - 1; c >= 0; c--) {
    var sheetNameSearch = sheets[c].getName();
    if (sheetNameSearch.indexOf("buffConsumables") > -1) {
      ss.deleteSheet(sheets[c]);
    }
  }
  var confSheet = confSpreadSheet.getSheetByName("buffConsumables");
  var conf = confSheet.copyTo(ss).hideSheet().setName("buffConsumables");

  var instructionsValues = []; instructionsValues[0] = []; instructionsValues[0].push(""); instructionsValues[1] = []; instructionsValues[1].push("");
  instructionsSheet.getRange(27, 2, 2, 1).setValues(instructionsValues);

  var darkMode = false;
  try {
    if (shiftRangeByRows(instructionsSheet, shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^export fights$").useRegularExpression(true).findNext(), 1), 2).getValue().indexOf("yes") > -1)
      darkMode = true;
  } catch { }

  sheet.getRange(firstPlayerNameRow - 2, firstPlayerNameColumn).clearContent();
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 70, 1).clearContent();
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 3, 70, 8).clearContent().clearNote().setFontStyle("normal").setFontLine("none");
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, 70, 10).setFontColor("black").setFontStyle("normal");
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 3, 70, 3).setFontWeight("normal").setFontSize(10);
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 6, 70, 4).setFontWeight("bold").setFontSize(10);

  if (darkMode) {
    sheet.getRange(1, 1, 80, 12).setBackground("#d9d9d9");
  } else {
    sheet.getRange(1, 1, 80, 12).setBackground("white");
  }
  sheet.getRange(6, 3, 1, 1).setBackground("#cccccc");

  var baseUrl = "https://fresh.warcraftlogs.com:443/v1/";
  if (lang != "EN") {
    baseUrl = "https://" + lang.toLowerCase() + ".fresh.warcraftlogs.com:443/v1/";
    baseUrlFrontEnd = "https://" + lang.toLowerCase() + ".fresh.warcraftlogs.com/reports/";
  }

  var api_key = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^2.$").useRegularExpression(true).findNext(), 4).getValue().replace(/\s/g, "");
  var reportPathOrId = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^3.$").useRegularExpression(true).findNext(), 4).getValue();
  var includeReportTitleInSheetNames = shiftRangeByColumns(instructionsSheet, instructionsSheet.createTextFinder("^4.$").useRegularExpression(true).findNext(), 4).getValue();
  var information = addColumnsToRange(sheet, addRowsToRange(sheet, sheet.createTextFinder("^" + getStringForLang("title", langKeys, langTrans, "", "", "", "") + " $").useRegularExpression(true).findNext(), 2), 1);
  addColumnsToRange(sheet, shiftRangeByColumns(sheet, information, 1), -1).clearContent();

  reportPathOrId = reportPathOrId.replace(".cn/", ".com/");
  var logId = "";
  if (reportPathOrId.indexOf("tbc.warcraftlogs") > -1)
    SpreadsheetApp.getUi().alert("This is the vanilla version of the CLA. Apparently you tried to run it for a TBC report. Please use the TBC version of the CLA for that, which you can get at https://discord.gg/nGvt5zH.");
  if (reportPathOrId.indexOf("classic.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("classic.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("vanilla.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("vanilla.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("sod.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("sod.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else if (reportPathOrId.indexOf("fresh.warcraftlogs.com/reports/") > -1)
    logId = reportPathOrId.split("fresh.warcraftlogs.com/reports/")[1].split("#")[0].split("?")[0];
  else
    logId = reportPathOrId;
  var apiKeyString = "?translate=true&api_key=" + api_key;

  var otherSheet = confSpreadSheet.getSheetByName("other");

  var bossesExcludedFromAllRaw = otherSheet.getRange(3, 9, 200, 1).getValues();
  var bossesExcludedFromAll = bossesExcludedFromAllRaw.reduce(function (ar, e) {
    if (e[0]) {
      ar.push(Number(e[0]))
    }
    return ar;
  }, []);

  var encounterIDFilter = "encounterid%20NOT%20IN%28";
  bossesExcludedFromAll.forEach(function (bossExcludedFromAll, bossExcludedFromAllCount) {
    encounterIDFilter = encounterIDFilter + bossExcludedFromAll.toString() + "%2C";
  })
  encounterIDFilter = encounterIDFilter.substr(0, encounterIDFilter.length - 3);
  encounterIDFilter = encounterIDFilter + "%29";

  var sectionToLookAtStart = 0;
  var sectionToLookAtEnd = 999999999999;
  var startEndString = "&start=" + sectionToLookAtStart.toString() + "&end=" + sectionToLookAtEnd.toString() + "&filter=" + encounterIDFilter;

  var manualStartAndEnd = "";
  try {
    var manualStartAndEnd = shiftRangeByColumns(sheet, sheet.createTextFinder(getStringForLang("startEndOptional", langKeys, langTrans, "", "", "", "")).findNext(), 1).getValue();
    if (manualStartAndEnd != null && manualStartAndEnd.toString().length > 0) {
      manualStartAndEnd = manualStartAndEnd.replace(new RegExp(' ', 'g'), "");
      var startEndParts = manualStartAndEnd.split("-");
      sectionToLookAtStart = startEndParts[0];
      sectionToLookAtEnd = startEndParts[1];
      startEndString = "&start=" + sectionToLookAtStart.toString() + "&end=" + sectionToLookAtEnd.toString();
    }
  } catch {
    var sectionToLookAtStart = 0;
    var sectionToLookAtEnd = 999999999999;
    var startEndString = "&start=" + sectionToLookAtStart.toString() + "&end=" + sectionToLookAtEnd.toString();
  }

  var allPlayersUrl = baseUrl + "report/tables/damage-taken/" + logId + apiKeyString + startEndString;
  var urlPeopleTracked = baseUrl + "report/tables/casts/" + logId + apiKeyString + startEndString;
  var urlAllFights = baseUrl + "report/fights/" + logId + apiKeyString + startEndString;
  var allFightsData = JSON.parse(UrlFetchApp.fetch(urlAllFights));

  var baseSheetName = getStringForLang("combatBuffsTab", langKeys, langTrans, "", "", "", "")
  if (includeReportTitleInSheetNames.indexOf("yes") > -1)
    baseSheetName += " " + allFightsData.title;
  try {
    sheet.setName(baseSheetName);
  } catch (err) {
    try {
      sheet.setName(baseSheetName + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", ""));
    } catch (err2) {
      try {
        sheet.setName(baseSheetName + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", ""));
      } catch (err3) {
        sheet.setName(baseSheetName + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", "") + "_" + getStringForLang("new", langKeys, langTrans, "", "", "", ""));
      }
    }
  }

  var informationValues = []; informationValues[0] = []; informationValues[1] = []; informationValues[2] = [];
  var returnVal = getRaidStartAndEnd(allFightsData, ss, baseUrl + "report/events/summary/" + logId + apiKeyString);
  var zonesFound = [];
  if (returnVal != null && returnVal.zonesFound != null)
    zonesFound = returnVal.zonesFound;
  var zoneTimesString = " (";
  if (zonesFound != null && zonesFound.length > 0) {
    zonesFound.forEach(function (raidZone, raidZoneCount) {
      zoneTimesString += raidZone[5] + " " + getStringForLang("in", langKeys, langTrans, "", "", "", "") + " ";
      if (raidZone[10] > 0) {
        zoneTimesString += getStringForTimeStamp(raidZone[10], true) + ", ";
      } else {
        zoneTimesString += getStringForTimeStamp(raidZone[2] - raidZone[1], true) + ", ";
      }
    })
    zoneTimesString = zoneTimesString.substr(0, zoneTimesString.length - 2);
    if (zoneTimesString.length > 0)
      informationValues[0].push(allFightsData.title + zoneTimesString + ")");
    else
      informationValues[0].push(allFightsData.title);
  } else
    SpreadsheetApp.getUi().alert(getStringForLang("noRaidZone", langKeys, langTrans, "", "", "", ""));

  allFightsData.fights.forEach(function (fight, fightCount) {
    if (fight.zoneName != null && fight.zoneName.length > 0 && informationValues[1].length == 0)
      informationValues[1].push(fight.zoneName);
  })
  if (allFightsData.zone <= 0) {
    SpreadsheetApp.getUi().alert(getStringForLang("zoneNotRecognized", langKeys, langTrans, "", "", "", ""));
  }

  var dateString = "";
  if (lang == "DE" || lang == "RU")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd.MM.yyyy HH:mm:ss");
  else if (lang == "EN")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "MMMM dd, yyyy HH:mm:ss");
  else if (lang == "CN")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "yyyy年M月d日 HH:mm:ss");
  else if (lang == "FR")
    dateString = Utilities.formatDate(new Date(allFightsData.start), "GMT+1", "dd/MM/yyyy HH:mm:ss");
  informationValues[2].push(dateString);
  sheet.getRange(information.getRow(), information.getColumn() + 1, 3, 1).setValues(informationValues);

  var spellIds = [];
  var spellIdStyles = [];

  var headersRaw = conf.getRange(1, 1, 1, 100).getValues();
  const headers = headersRaw.filter(function (x) {
    return !(x.every(element => element === (undefined || null || '')))
  });
  var headerDoneCount = 0;
  headers[0].forEach(function (header, headerCount) {
    if (header != null && header.length != null && header.length > 0) {
      spellIds[headerDoneCount] = getEntriesForHeader(headers[0], header, conf);
      spellIdStyles[headerDoneCount] = getStylesForHeader(headers[0], header, conf);
      headerDoneCount++;
    }
  })

  var bossesTimeSpans = [];
  var numberOfTotalBossFights = 0;
  var bossFightsExcluded = "";
  var previousFightStartAll = -1;
  var previousFightEndAll = -1;
  var previousOriginalBossAll = -1;
  var previousFightId = -1;
  var bossFightsWithoutCombatantInfo = [];
  var bossFightCount = 0;
  var bossSummaryDataAll = [];
  var bossDamageDataAll = [];
  allFightsData.fights.forEach(function (bossFight, bossFightRawCount) {
    if (bossFightCount % 10 == 0)
      Utilities.sleep(300);
    if (bossFight.boss != null && bossFight.boss > 0) {
      bossFightCount++;
      var urlAllCombatantInfos = baseUrl + "report/events/summary/" + logId + apiKeyString + "&start=" + previousFightStartAll + "&end=" + bossFight.end_time + "&filter=type%20%3D%20%22combatantinfo%22";
      var allCombatantInfoData = JSON.parse(UrlFetchApp.fetch(urlAllCombatantInfos));
      var isBossExcluded = false;
      bossesExcludedFromAll.forEach(function (bossExcludedFromAll, bossExcludedFromAllCount) {
        if (bossFight.start_time >= sectionToLookAtStart && bossFight.end_time <= sectionToLookAtEnd && (bossExcludedFromAll == bossFight.boss || "10" + bossExcludedFromAll == bossFight.boss)) {
          isBossExcluded = true;
          if (bossFightsExcluded.indexOf(bossFight.name) < 0)
            bossFightsExcluded += bossFight.name + " & ";
          else {
            var previousNumber = 1;
            var temp = bossFightsExcluded.split("x " + bossFight.name)[0];
            if (temp.length < 4) {
              previousNumber = Number(temp);
              var stringOld = previousNumber.toString() + "x " + bossFight.name;
            } else if (temp.length == bossFightsExcluded.length) {
              var stringOld = bossFight.name;
            } else {
              var tempArr = temp.split(" & ");
              previousNumber = Number(tempArr[tempArr.length - 1]);
              var stringOld = previousNumber.toString() + "x " + bossFight.name;
            }
            previousNumber++;
            var stringNew = previousNumber.toString() + "x " + bossFight.name;
            bossFightsExcluded = bossFightsExcluded.replace(stringOld, stringNew);
          }
        }
      })
      if (bossFight.start_time >= sectionToLookAtStart && bossFight.end_time <= sectionToLookAtEnd && !isBossExcluded && ((bossFight.end_time - bossFight.start_time) < 100 || ((bossFight.end_time - bossFight.start_time) > 3000))) {
        var combatantInfoDataFound = false;
        allCombatantInfoData.events.forEach(function (combatantInfoData, combatantInfoDataCount) {
          if (combatantInfoData.fight != null && (combatantInfoData.fight == bossFight.id || (combatantInfoData.fight == previousFightId && (previousOriginalBossAll == bossFight.boss || "10" + previousOriginalBossAll.toString() == bossFight.boss.toString()) && ((previousFightEndAll - previousFightStartAll) < 100 || ((previousFightEndAll - previousFightStartAll) > 3000)))))
            combatantInfoDataFound = true;
        })
        if (combatantInfoDataFound) {
          if ((previousOriginalBossAll == bossFight.boss || "10" + previousOriginalBossAll.toString() == bossFight.boss.toString()) && ((previousFightEndAll - previousFightStartAll) < 100 || ((previousFightEndAll - previousFightStartAll) > 3000))) {
            var timeSpanArr = [];
            timeSpanArr.push(previousFightStartAll);
            timeSpanArr.push(previousFightEndAll);
            timeSpanArr.push(bossFight.start_time);
            timeSpanArr.push(bossFight.id);
            timeSpanArr.push(bossFight.name);
            bossesTimeSpans.push(timeSpanArr);
          }
          var timeSpanArr = [];
          timeSpanArr.push(bossFight.start_time);
          timeSpanArr.push(bossFight.end_time);
          timeSpanArr.push(bossFight.start_time);
          timeSpanArr.push(bossFight.id);
          timeSpanArr.push(bossFight.name);
          bossesTimeSpans.push(timeSpanArr);
          bossSummaryDataAll[numberOfTotalBossFights] = [];
          bossSummaryDataAll[numberOfTotalBossFights].push(JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace(startEndString, "&start=" + bossFight.start_time + "&end=" + bossFight.end_time).replace("damage-taken", "summary"))));
          bossSummaryDataAll[numberOfTotalBossFights].push(bossFight.id.toString());
          bossDamageDataAll[numberOfTotalBossFights] = [];
          bossDamageDataAll[numberOfTotalBossFights].push(JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace(startEndString, "&start=" + bossFight.start_time + "&end=" + bossFight.end_time).replace("damage-taken", "damage-done") + "&abilityid=11357")));
          bossDamageDataAll[numberOfTotalBossFights].push(bossFight.id.toString());
          numberOfTotalBossFights++;
        } else {
          if (bossFightsExcluded.indexOf(bossFight.name) < 0)
            bossFightsExcluded += bossFight.name + " & ";
          else {
            var previousNumber = 1;
            var temp = bossFightsExcluded.split("x " + bossFight.name)[0];
            if (temp.length < 4) {
              previousNumber = Number(temp);
              var stringOld = previousNumber.toString() + "x " + bossFight.name;
            } else if (temp.length == bossFightsExcluded.length) {
              var stringOld = bossFight.name;
            } else {
              var tempArr = temp.split(" & ");
              previousNumber = Number(tempArr[tempArr.length - 1]);
              var stringOld = previousNumber.toString() + "x " + bossFight.name;
            }
            previousNumber++;
            var stringNew = previousNumber.toString() + "x " + bossFight.name;
            bossFightsExcluded = bossFightsExcluded.replace(stringOld, stringNew);
          }
          if (bossFightsWithoutCombatantInfo.indexOf(bossFight.id) < 0) {
            bossFightsWithoutCombatantInfo.push(bossFight.id);
          }
        }
      }
    }
    previousFightStartAll = bossFight.start_time;
    previousFightEndAll = bossFight.end_time;
    previousFightId = bossFight.id;
    if (bossFight.originalBoss != null)
      previousOriginalBossAll = bossFight.originalBoss;
    else
      previousOriginalBossAll = -1;
  })
  var spacer = 2;
  var stringToPrint = "";
  if (bossFightsExcluded.length > 2)
    stringToPrint = getStringForLang("nrBossFights", langKeys, langTrans, "", "", "", "") + numberOfTotalBossFights.toString() + "\r\n(" + bossFightsExcluded.substring(0, bossFightsExcluded.length - 3) + " " + getStringForLang("removed", langKeys, langTrans, "", "", "", "") + ")";
  else
    stringToPrint = getStringForLang("nrBossFights", langKeys, langTrans, "", "", "", "") + numberOfTotalBossFights.toString();
  if (Math.ceil(stringToPrint.length / 22) - 2 > spacer)
    spacer = Math.ceil(stringToPrint.length / 22) - 2;
  sheet.getRange(firstPlayerNameRow - 2, firstPlayerNameColumn).setValue(stringToPrint);

  var allPlayersData = JSON.parse(UrlFetchApp.fetch(urlPeopleTracked));

  const allPlayersByNameAsc = sortByProperty(sortByProperty(allPlayersData.entries, 'name'), "type");

  var playersFound = 0;
  var playerRows = [];
  var playerNames = [];
  var playerColours = [];
  var richTexts = [];
  var headerValues = sheet.getRange(firstPlayerNameRow - 2, firstPlayerNameColumn + 3, 1, spellIds.length).getValues();
  allPlayersByNameAsc.forEach(function (playerByNameAsc, playerCountByNameAsc) {
    if ((playerByNameAsc.type == "Druid" || playerByNameAsc.type == "Hunter" || playerByNameAsc.type == "Mage" || playerByNameAsc.type == "Priest" || playerByNameAsc.type == "Paladin" || playerByNameAsc.type == "Rogue" || playerByNameAsc.type == "Shaman" || playerByNameAsc.type == "Warlock" || playerByNameAsc.type == "Warrior") && playerByNameAsc.total > 20) {
      var playerFightCount = [];
      var buffData = JSON.parse(UrlFetchApp.fetch(allPlayersUrl.replace("damage-taken", "buffs") + "&sourceid=" + playerByNameAsc.id));
      var bossCovered = [];
      var badBuffsFound = [];

      allFightsData.fights.forEach(function (bossFight, bossFightRawCount) {
        allFightsData.friendlies.forEach(function (friendlyPlayer, friendlyPlayerCount) {
          if (bossFight.boss != null && bossFight.boss > 0 && bossFight.start_time >= sectionToLookAtStart && bossFight.end_time <= sectionToLookAtEnd) {
            if (friendlyPlayer.id == playerByNameAsc.id) {
              friendlyPlayer.fights.forEach(function (playerFight, playerFightRawCount) {
                var isBossExcluded = false;
                bossesExcludedFromAll.forEach(function (bossExcludedFromAll, bossExcludedFromAllCount) {
                  if (bossFight.boss != null && bossFight.boss > 0 && bossFight.start_time >= sectionToLookAtStart && bossFight.end_time <= sectionToLookAtEnd && (bossExcludedFromAll == bossFight.boss || "10" + bossExcludedFromAll == bossFight.boss))
                    isBossExcluded = true;
                })
                bossFightsWithoutCombatantInfo.forEach(function (bossFightWithoutCombatantInfo, bossFightWithoutCombatantInfoCount) {
                  if (bossFightWithoutCombatantInfo == bossFight.id)
                    isBossExcluded = true;
                })
                if (bossFight.id == playerFight.id && !isBossExcluded) {
                  playerFightCount.push(bossFight.id);
                }
              })
            }
          }
        })
      })

      if (playerFightCount.length > 0) {
        richTexts[playersFound] = [];
        playerRows[playersFound] = [];
        playerColours[playersFound] = [];
        playerColours[playersFound].push(getColourForPlayerClass(playerByNameAsc.type));
        playerNames[playersFound] = [];
        playerNames[playersFound].push(playerByNameAsc.name);

        spellIds.forEach(function (spellIdsColumn, spellIdsColumnCount) {
          var headerValue = headerValues[0][spellIdsColumnCount];
          if (headerValue != getStringForLang("weaponEnhancement", langKeys, langTrans, "", "", "", "")) {
            var bossesFound = [];
            var isPlayerACheapass = false;
            spellIdsColumn.forEach(function (spellIdsCell, spellIdsCellCount) {
              buffData.auras.forEach(function (buff, buffCount) {
                if (buff.guid != null && buff.guid.toString().length > 0 && buff.guid.toString() == spellIdsCell.toString().split(" [")[0]) {
                  buff.bands.forEach(function (buffBand, buffBandCount) {
                    bossesTimeSpans.forEach(function (bossTimeSpan, bossTimeSpanCount) {
                      var swiftnessZanzaUsed = false;
                      if ((headerValue != getStringForLang("flaskText", langKeys, langTrans, "", "", "", "") || (headerValue == getStringForLang("flaskText", langKeys, langTrans, "", "", "", "") && bossCovered.indexOf(bossTimeSpan[2].toString()) < 0))) {
                        if (headerValue == getStringForLang("battleElixir", langKeys, langTrans, "", "", "", "")) {
                          buffData.auras.forEach(function (buffToo, buffTooCount) {
                            buffToo.bands.forEach(function (buffBandToo, buffBandTooCount) {
                              if (buffToo.guid != null && buffToo.guid.toString().length > 0 && buffToo.guid.toString() == "24383" && ((buffBandToo.endTime >= bossTimeSpan[0] && buffBandToo.endTime <= bossTimeSpan[1]) || (buffBandToo.startTime >= bossTimeSpan[0] && buffBandToo.endTime <= bossTimeSpan[1]))) {
                                swiftnessZanzaUsed = true;
                              }
                            })
                          })
                        }
                        if (bossesFound.indexOf(bossTimeSpan[2].toString()) < 0 && ((buffBand.endTime >= bossTimeSpan[0] && buffBand.endTime <= bossTimeSpan[1]) || (buffBand.startTime >= bossTimeSpan[0] && buffBand.endTime <= bossTimeSpan[1]))) {
                          bossesFound.push(bossTimeSpan[2].toString());
                          if (headerValue == getStringForLang("battleElixir", langKeys, langTrans, "", "", "", "") || headerValue == getStringForLang("guardianElixir", langKeys, langTrans, "", "", "", "")) {
                            bossCovered.push(bossTimeSpan[2].toString());
                          }
                          try {
                            if ((spellIdStyles[spellIdsColumnCount][spellIdsCellCount] == "italic") && !swiftnessZanzaUsed) {
                              isPlayerACheapass = true;
                              var badBuffAlreadyFound = false;
                              badBuffsFound.forEach(function (badBuff, badBuffCount) {
                                if (badBuff[0] == buff.guid.toString())
                                  badBuffAlreadyFound = true;
                              })
                              if (!badBuffAlreadyFound) {
                                var badBuffFound = [];
                                badBuffFound.push(buff.guid.toString());
                                badBuffFound.push(buff.name.toString());
                                badBuffsFound.push(badBuffFound);
                              }
                            }
                          } catch { }
                        }
                      }
                    })
                  })
                }
              })
            })
            if (bossesFound != null && bossesFound.length > 0) {
              if (isPlayerACheapass) {
                var richText = SpreadsheetApp.newRichTextValue();
                var textStyle = SpreadsheetApp.newTextStyle();
                textStyle.setForegroundColor("#9a86d6"); textStyle.setFontSize(1);
                richText.setText(Math.round(bossesFound.length * 100 / playerFightCount.length) + "%*");
                richText.setTextStyle((Math.round(bossesFound.length * 100 / playerFightCount.length) + "%*").length - 1, (Math.round(bossesFound.length * 100 / playerFightCount.length) + "%*").length, textStyle.build());
                playerRows[playersFound].push(richText.build());
              } else {
                playerRows[playersFound].push(SpreadsheetApp.newRichTextValue().setText(Math.round(bossesFound.length * 100 / playerFightCount.length) + "%").build());
              }
            } else {
              playerRows[playersFound].push(SpreadsheetApp.newRichTextValue().setText("0%").build());
            }
          } else {
            var usedTemporaryWeaponEnchant = 0;
            var atLeastOneConsecrationItem = "no";
            var isPlayerACheapass = false;
            var found = 0;
            var total = 0;
            bossSummaryDataAll.forEach(function (bossSummaryData, bossSummaryDataCount) {
              var increaseTotal = false;
              var increaseFound = false;
              if (increaseTotal == false && bossSummaryData[0].playerDetails != null && bossSummaryData[0].playerDetails.dps != null && bossSummaryData[0].playerDetails.dps.length > 0) {
                bossSummaryData[0].playerDetails.dps.forEach(function (playerInfo, playerInfoCount) {
                  if (playerInfo.name == playerByNameAsc.name) {
                    if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                      increaseTotal = true;
                      playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                        if (item.id != null && item.id.toString().length > 0) {
                          if (item.slot != null && item.slot.toString() == "15" || item.slot.toString() == "16") {
                            if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                              increaseTotal = false;
                            } else {
                              if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266") {
                                if (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")
                                  atLeastOneConsecrationItem = "yes";
                                increaseFound = true;
                              }
                            }
                          }
                        }
                      })
                    }
                  }
                })
              }
              if (increaseTotal == false && bossSummaryData[0].playerDetails != null && bossSummaryData[0].playerDetails.healers != null && bossSummaryData[0].playerDetails.healers.length > 0) {
                bossSummaryData[0].playerDetails.healers.forEach(function (playerInfo, playerInfoCount) {
                  if (playerInfo.name == playerByNameAsc.name) {
                    if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                      increaseTotal = true;
                      playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                        if (item.id != null && item.id.toString().length > 0) {
                          if (item.slot != null && item.slot.toString() == "15" || item.slot.toString() == "16") {
                            if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                              increaseTotal = false;
                            } else {
                              if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266") {
                                if (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")
                                  atLeastOneConsecrationItem = "yes";
                                increaseFound = true;
                              }
                            }
                          }
                        }
                      })
                    }
                  }
                })
              }
              if (increaseTotal == false && bossSummaryData[0].playerDetails != null && bossSummaryData[0].playerDetails.tanks != null && bossSummaryData[0].playerDetails.tanks.length > 0) {
                bossSummaryData[0].playerDetails.tanks.forEach(function (playerInfo, playerInfoCount) {
                  if (playerInfo.name == playerByNameAsc.name) {
                    if (playerInfo.combatantInfo != null && playerInfo.combatantInfo.gear != null) {
                      increaseTotal = true;
                      playerInfo.combatantInfo.gear.forEach(function (item, itemCount) {
                        if (item.id != null && item.id.toString().length > 0) {
                          if (item.slot != null && item.slot.toString() == "15" || item.slot.toString() == "16") {
                            if (item.id.toString() == "19022" || item.id.toString() == "19970" || item.id.toString() == "25978" || item.id.toString() == "6365" || item.id.toString() == "12225" || item.id.toString() == "6367" || item.id.toString() == "6366" || item.id.toString() == "6256") {
                              increaseTotal = false;
                            } else {
                              if (item.temporaryEnchantName != null && item.temporaryEnchantName.length != null && item.temporaryEnchantName.length > 0 && item.temporaryEnchant.toString() != "4264" && item.temporaryEnchant.toString() != "263" && item.temporaryEnchant.toString() != "264" && item.temporaryEnchant.toString() != "265" && item.temporaryEnchant.toString() != "266") {
                                if (item.temporaryEnchant.toString() == "2684" || item.temporaryEnchant.toString() == "2685")
                                  atLeastOneConsecrationItem = "yes";
                                increaseFound = true;
                              }
                            }
                          }
                        }
                      })
                    }
                  }
                })
              }
              if (increaseFound)
                found++;
              else if (playerByNameAsc.type == "Rogue" && increaseTotal) {
                bossDamageDataAll.forEach(function (bossDamageData, bossDamageDataCount) {
                  if (bossDamageData[1] == bossSummaryData[1]) {
                    bossDamageData[0].entries.forEach(function (bossDamageData, bossDamageDataCount) {
                      if (bossDamageData != null && bossDamageData.id != null && bossDamageData.id == playerByNameAsc.id) {
                        found++;
                      }
                    })
                  }
                })
              }
              if (increaseTotal)
                total++;
            })
            usedTemporaryWeaponEnchant = Math.round(found * 100 / total) + "%";

            if (total > 0) {
              if (isPlayerACheapass) {
                var richText = SpreadsheetApp.newRichTextValue();
                var textStyle = SpreadsheetApp.newTextStyle();
                textStyle.setForegroundColor("#9a86d6"); textStyle.setFontSize(1);
                richText.setText(usedTemporaryWeaponEnchant + "*");
                richText.setTextStyle((usedTemporaryWeaponEnchant + "*").length - 1, (usedTemporaryWeaponEnchant + "*").length, textStyle.build());
                playerRows[playersFound].push(richText.build());
              } else {
                if (atLeastOneConsecrationItem == "yes") {
                  sheet.getRange(firstPlayerNameRow + playersFound, firstPlayerNameColumn + spellIdsColumnCount + 2).setBackground("#666666");
                }
                playerRows[playersFound].push(SpreadsheetApp.newRichTextValue().setText(usedTemporaryWeaponEnchant).build());
              }
            } else {
              playerRows[playersFound].push(SpreadsheetApp.newRichTextValue().setText("0%").build());
            }
          }
        })
        if (badBuffsFound.length > 0) {
          var suboptimalStuffFoundString = "";
          badBuffsFound.forEach(function (badBuffFound, badBuffFoundCount) {
            suboptimalStuffFoundString += badBuffFound[1] + ", ";
          })
          var richValue = SpreadsheetApp.newRichTextValue().setText(suboptimalStuffFoundString.substr(0, suboptimalStuffFoundString.length - 2));
          var wowheadBaseUrl = "https://wowhead.com/classic/spell=";
          if (lang != "EN")
            wowheadBaseUrl = wowheadBaseUrl.replace("wotlk/", "wotlk/" + lang.toLowerCase() + "/");
          var currentPosition = 0;
          badBuffsFound.forEach(function (badBuffFound, badBuffFoundCount) {
            richValue.setLinkUrl(currentPosition, currentPosition + (badBuffFound[1].length), wowheadBaseUrl + badBuffFound[0]);
            currentPosition += (badBuffFound[1].length + 2);
          })
          richTexts[playersFound].push(richValue.build());
        } else {
          richTexts[playersFound].push(SpreadsheetApp.newRichTextValue().setText("").build());
        }

        playersFound++;
      }
    }
  })
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 3 + spellIds.length, playersFound, 1).setRichTextValues(richTexts);
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn, playersFound, 1).setValues(playerNames).setBackgrounds(playerColours).setFontWeight("bold");
  sheet.getRange(firstPlayerNameRow, firstPlayerNameColumn + 3, playersFound, spellIds.length).setRichTextValues(playerRows);
  try {
    var sheets = ss.getSheets();
    for (var c = sheets.length - 1; c >= 0; c--) {
      var sheetNameSearch = sheets[c].getName();
      if (sheetNameSearch.indexOf("buffConsumables") > -1) {
        ss.deleteSheet(sheets[c]);
      }
    }
  } catch { }
}

function getEntriesForHeader(headers, headerToSearch, conf) {
  var columnNumber = -1;
  headers.forEach(function (header, headerCount) {
    if (header == headerToSearch)
      columnNumber = headerCount;
  })
  var entriesRaw = conf.getRange(2, columnNumber + 1, 200, 1).getValues();
  return entriesRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);
}

function getStylesForHeader(headers, headerToSearch, conf) {
  var columnNumber = -1;
  headers.forEach(function (header, headerCount) {
    if (header == headerToSearch)
      columnNumber = headerCount;
  })
  var entriesRaw = conf.getRange(2, columnNumber + 1, 200, 1).getFontStyles();
  return entriesRaw.reduce(function (ar, e) {
    if (e[0]) ar.push(e[0])
    return ar;
  }, []);
}
