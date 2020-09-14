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

	let html = `<div id="individualSheetErrors" class="errorsInfo"><h2 class="errorHeader">🛑 Error 🛑</h2><br>`;
	
	let missingNames = checkForMissingNames(names, sheets);
	if (missingNames.length > 0) {
		html += `<h2 class="errorHeader2">The following are missing individual sheets:</h2><br>`;
		for (const name of missingNames) {
			html += `<p class="errorName">${name}</p>`;
		}
		html += `<br><h3 class="errorTip">If this error is still coming up while all sheets exist, make sure the spelling is correct in the sheet name, as well as making sure the individual sheet's formatting is correct. For example, make sure the data starts at A1 (A1 should be the "Address" header)</h3>`;
	}
	html += `</div>`;

	html += `<div id="secondCheckErrors" class="errorsInfo">`;

	let secondCheckOutput = checkBigSheet(fileID);

	html += secondCheckOutput;

	html += `</div>`;


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

	let output = ``;
	for (const name in incorrectNames) {
		output += `${name} made ${incorrectNames[name].actual} calls, but only ${incorrectNames[name].counted} was counted.<br>`;
	}

	return output;
}


function startReport(teams, params, fileID) {
	let oss = SpreadsheetApp.openById(fileID);
	let sheets = oss.getSheets();

	let nameList = [];
	for (const team in teams) {
		for (const name of teams[team]) {
			nameList.push(name);
		}
	}

	let output = checkSheetForErrors(teams, nameList, fileID);
	return output;
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