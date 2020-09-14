function getNamesVAN(fileID, fileName) {
	try {
		let newFile = convertToSheet(fileID);
		let names = readNames(newFile.id);
		
		for (const name of names) {
			if (isNaN(name) == false || name == "Address") {
				return [0,0,'Error'];
			}
		}
		return [names, newFile.id];
	} catch(err) {
		return [0,0,'Error'];
	}
}

function readNames(fileID) {
	let ss = SpreadsheetApp.openById(fileID);
	let overviewSheet = ss.getSheets()[0];

	let firstColumnRange = overviewSheet.getRange("A1").getDataRegion(SpreadsheetApp.Dimension.ROWS);
	return firstColumnRange.getValues();
}


function convertToSheet(fileID) {
	let excelFile = DriveApp.getFileById(fileID);
	let parentFolder = excelFile.getParents().next();

	let blob = excelFile.getBlob();

	let resource = {
		title: `${excelFile.getName()} CONVERTED`,
		mimeType: MimeType.GOOGLE_SHEETS,
		parents: [{id: parentFolder.getId()}],
	};

	return Drive.Files.insert(resource, blob);
}

function checkSheetForErrors(teams, names, fileID) {
	let ss = SpreadsheetApp.openById(fileID);
	let sheets = ss.getSheets();
	let html = "";
	
	let missingNames = checkForMissingNames(names, sheets);
	if (missingNames.length > 0) {
		html += `<div id="individualSheetErrors" class="errorsInfo"><h2 class="errorHeader">🛑 Error 🛑</h2><br>`;
		html += `<h2 class="errorHeader2">The following are missing individual sheets:</h2><br>`;
		for (const name of missingNames) {
			html += `<p class="errorName">${name}</p>`;
		}
		html += `<br><h3 class="errorTip">If this error is still coming up while all sheets exist, make sure the spelling is correct in the sheet name, as well as making sure the individual sheet's formatting is correct. For example, make sure the data starts at A1 (A1 should be the "Address" header)</h3>`;
		html += `</div>`;
	}
	let secondCheckOutput = checkBigSheet(fileID);
	html += secondCheckOutput;
	
	return html;
}

function checkBigSheet(fileID) {
	let ss = SpreadsheetApp.openById(fileID);
	let overviewSheet = ss.getSheets()[0];
	let bigSheet = ss.getSheets()[1];
	
	let nameCounts = {};
	let checkCounts = {};
	
	let names = overviewSheet.getRange("A1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let counts = overviewSheet.getRange("C1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	
	for (let i = 0; i < names.length; i++) {
		nameCounts[names[i]] = counts[i];
		checkCounts[names[i]] = 0;
	}
	
	let bigSheetVals = bigSheet.getRange('F1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	
	if (bigSheetVals[0] == 'Canvassed By') {
		bigSheetVals.shift()
	}
	
	try {
		for (const name of bigSheetVals) {
			checkCounts[name]++;
		}
	} catch(err) {
		let output = `<h2 class="errorHeader">🛑 Error 🛑</h2><br>`;
		output += `<h2 class="errorHeader2">There is an unknown error in the big sheet.</h2><br>`;
		output += `<h3 class="errorTip">Make sure the data starts on A1 and make sure there are no empty rows. You can do this by selecting the first cell, and then pressing <b>Ctrl + Down</b>. This will bring you to the next empty cell, which should be the last cell.</h3>`;
		return output;
	}
	
	let incorrectNames = {};
	
	for (const name in nameCounts) {
		let real = nameCounts[name];
		let big = checkCounts[name];
		
		if (real != big) {
			incorrectNames[name] = {
				actual: real,
				counted: big,
			};
		}
	}
	
	if (Object.keys(incorrectNames).length > 0) {
		let output = `<div id="secondCheckErrors" class="errorsInfo">`;
			output += `<h2 class="errorHeader">🛑 Error 🛑</h2><br>`;
			output += `<h2 class="errorHeader2">There are discrepancies in the big sheet.</h2><br>`;
		for (const name in incorrectNames) {
			output += `<p class="discrepancyInfo">${name} made <b>${incorrectNames[name].actual}</b> calls, but only <b>${incorrectNames[name].counted}</b> were counted.</p><br>`;
		}
		output += `<h3 class="errorTip">Double check the big sheet and make sure the number of records in the big sheet per canvasser matches the overview number of call attempts.</h3>`
		output += `</div>`;
		return output;
	}
	
	return '';
}


function startReport(teams, params, fileID) {
	let oss = SpreadsheetApp.openById(fileID);
	let sheets = oss.getSheets();

	let fileName = DriveApp.getFileById(fileID).getName();
	let dateStr = fileName.slice(0,6);
	dateStr = `${dateStr.slice(0,2)}-${dateStr.slice(2,4)}-${dateStr.slice(4,6)}`;

	let parentFolder = DriveApp.getFileById(fileID).getParents().next();
	let newFolder = parentFolder.createFolder(`${dateStr} GENERATED`);


	let nameList = [];
	for (const team in teams) {
		for (const name of teams[team]) {
			nameList.push(name);
		}
	}

	let output = checkSheetForErrors(teams, nameList, fileID);
	if (output != "") {
		return output;
	}

	let overviewData = extractOverviewData(fileID);
	let timeInfo = getTimeInfo(nameList, fileID);

	for (const name in overviewData) {
		overviewData[name]['timeInfo'] = timeInfo[name];
	}

	for (const team in teams) {
		Logger.log(`creating ss for ${team}`);
		let teamSSID = createSpreadsheetNamed(newFolder, `${dateStr} TEAM ${team}`);
		Logger.log(`ssid: ${teamSSID}`);
		createOverviewPage(overviewData, teamSSID);
	}
}



function createOverviewPage(data, ssid) {
	console.log(`renamining sheet for ${ssid}`);
	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Testing 123');
	
}

function getTimeInfo(names, fileID) {
	let ss = SpreadsheetApp.openById(fileID);
	let sheets = {};
	let timeInfo = {};
	
	for (const sheet of ss.getSheets()) {
		sheets[sheet.getName()] = sheet;
	}

	for (const name of names) {
		let info = getIndividualTimeInfo(name, sheets[name]);
		timeInfo[name] = {
			startTime: info[0],
			endTime: info[1],
			timeDiffs: info[2],
			avgTimeDiff: info[3],
		};
	}

	return timeInfo;
}

function getIndividualTimeInfo(name, sheet) {
	let timeCol = sheet.getRange('C1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	timeCol.shift();

	let startTime = timeCol[0];
	let endTime = timeCol[timeCol.length-1];
	let timeDiffs = getTimeDiffs(timeCol);
	let avgTimeDiff = averageTimeDiffs(timeDiffs);

	return [startTime, endTime, timeDiffs, avgTimeDiff];
}

function averageTimeDiffs(diffs) {
	let avg = 0;
	for (const diff of diffs) {
		avg += diff;
	}
	return avg / diffs.length;
}

function getTimeDiffs(timeCol) {
	let diffs = [];

	for (let i = 0; i < timeCol.length - 1; i++) {
		let boilerPlate = 'January 1, 2020 ';
		let firstTime = new Date(timeCol[i]);
		let secondTime = new Date(timeCol[i+1]);

		let diff = (secondTime - firstTime) / 60000;
		diffs.push(diff);
	}

	return diffs;
}

function extractOverviewData(fileID) {
	// returns obj with name as key
	// and each relevant piece of data
	let ss = SpreadsheetApp.openById(fileID);
	let sheet = ss.getSheets()[0];

	let outputData = {};

	let nameCol = sheet.getRange('A1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let totalCol = sheet.getRange('C1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let canvCol = sheet.getRange('D1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let leftMessageCol = sheet.getRange('E1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let refusedCol = sheet.getRange('F1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let otherLanguageCol = sheet.getRange('G1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let notHomeCol = sheet.getRange('J1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let disconnectedCol = sheet.getRange('K1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let movedCol = sheet.getRange('L1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let otherCol = sheet.getRange('M1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();

	let i;

	for (i = 0; i < nameCol.length; i++) {
		outputData[nameCol[i]] = {
			totalCalls: totalCol[i],
			canvassed: canvCol[i],
			leftMessage: leftMessageCol[i],
			refused: refusedCol[i],
			otherLanguage: otherLanguageCol[i],
			notHome: notHomeCol[i],
			disconnected: disconnectedCol[i],
			moved: movedCol[i],
			other: otherCol[i],
		};
	}

	return outputData;
}

function checkForMissingNames(names, sheets) {
	let missing = [];
	for (const name of names) {
		if (!checkNameInSheets(name, sheets)) {
			missing.push(name);
		}
	}

	return missing;
}

function checkNameInSheets(name, sheets) {
	let i;
	for (i = 0; i < sheets.length; i++) {
		if (name == sheets[i].getName()) {
			return true;
		}
	}
	return false;
}