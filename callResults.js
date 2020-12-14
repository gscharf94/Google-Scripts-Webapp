function callResultsStart(fileID, fileName) {
	log(`starting callResultsStart() fileID=${fileID} fileName=${fileName}`);
	let file = DriveApp.getFileById(fileID);
	let parentFolder = file.getParents().next();
	let blob = file.getBlob();
	let rawText = blob.getDataAsString('utf8');
	log(`succesfully loaded data`);

	let ssID = createSpreadsheet(fileName, parentFolder, 'CR-');
	let parsedCSV = parseCSV(rawText);
	let data = generateDataDict(parsedCSV);

	let timeInfo = createIndividualSheetsCallResults(ssID, data);
	createOverviewSheetCallResults(ssID, data, timeInfo);

	return SpreadsheetApp.openById(ssID).getUrl();
}

function createIndividualSheetsCallResults(ssID, data) {
	log(`starting individual sheets ssID=${ssID}`);

	// returns timeinfo for overview page
	let ss = SpreadsheetApp.openById(ssID);

	let timeInfo = {};
	for (const name in data) {
		let timeArr = createIndividualSheetCallResults(name, data[name], ss);
		timeInfo[name] = {};
		timeInfo[name]['timeDiffs'] = timeArr[0];
		timeInfo[name]['startTime'] = timeArr[1];
		timeInfo[name]['endTime'] = timeArr[2];
	}

	return timeInfo;
}

function createOverviewSheetCallResults(ssID, data, timeInfo) {
	let ss = SpreadsheetApp.openById(ssID);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Overview');

	let resultsTemplate = createResultsTemplate(data);

	let topRow = ['Caller ID', 'Total Calls'];
	for (const result in resultsTemplate) {
		topRow.push(result);
		topRow.push(`${result.slice(0,4)} Avg`);
	}
	topRow.push('Start Time');
	topRow.push('End Time');
	topRow.push('Hours Worked');
	topRow.push('5+ Time Diffs');

	let topRowRange = addRange(ssID, topRow, 1, 1, 'Overview', 'row');
	let topRowRange2 = sheet.getRange(`B1:${LETTERS[topRow.length]}1`);
	topRowRange2.setFontSize(11);

	let mainRangeValues = [];
	for (const name in data) {
		let individualResults = fillInTemplate(data[name], resultsTemplate);
		let row = [name, data[name].length];
		for (const result in individualResults) {
			row.push(individualResults[result]);
			row.push(individualResults[result] / data[name].length);
		}
		mainRangeValues.push(row);
	}

	let mainRange = addRange(ssID, mainRangeValues, 2, 1, 'Overview');
	addAverageRow(sheet, resultsTemplate, mainRangeValues);
	formatOverviewSheetCallResults(sheet, topRow.length, mainRangeValues.length, resultsTemplate, mainRange);
	addTimeInfo(sheet, timeInfo, data, resultsTemplate, ssID);

}

function countTimeDiffs(name, data) {
	const THRESHOLD = 5;
	let count = 0;

	data.forEach((val) => {
		let num = Number(String(val).split(" ")[0]);
		if (num >= THRESHOLD) {
			count++;
		}
	});
	return count;
}


function addTimeInfo(sheet, timeInfo, data, template, ssID) {
	let mainRangeValues = [];
	for (const name in data) {
		let sTime = timeInfo[name]['startTime'];
		let eTime = timeInfo[name]['endTime'];
		let hWorked = new Date('1970/01/01 ' + eTime) - new Date('1970/01/01 ' + sTime);
		hWorked = Math.round(hWorked / 3600000 * 10) / 10;
		let tDiffCount = countTimeDiffs(name, timeInfo[name]['timeDiffs']);

		let row = [
			sTime,
			eTime,
			hWorked,
			tDiffCount,
		];
		mainRangeValues.push(row);
	}

	let timeRange = addRange(ssID, mainRangeValues, 2, 9, 'Overview');

}

function formatOverviewSheetCallResults(sheet, width, height, template, mainRange) {

	sheet.setHiddenGridlines(true);

	sheet.setColumnWidth(1, 143);
	let i;
	for (i = 2; i < width + 1; i++) {
		sheet.setColumnWidth(i, 73);
	}
	sheet.setRowHeight(1, 57);

	let col = 4;
	for (const result in template) {
		col += 2;
	}

	sheet.setColumnWidth(col - 1, 92);
	sheet.setColumnWidth(col, 92);
	sheet.setColumnWidth(col + 1, 61);
	sheet.setColumnWidth(col + 2, 61);


	let topRowRange = sheet.getRange(`A1:${LETTERS[width]}1`);
	let topRowRange2 = sheet.getRange(`B1:${LETTERS[width]}1`);
	topRowRange.setFontWeight('bold');
	topRowRange.setHorizontalAlignment('center');
	topRowRange.setWrap(true);

	topRowRange2.setBackground(COLORS['darkGray']);

	let leftColRange = sheet.getRange(`A1:A${height+3}`);
	leftColRange.setFontWeight('bold');
	leftColRange.setHorizontalAlignment('right');

	let leftRange = sheet.getRange(`A2:A${height+3}`);
	let nextRange = sheet.getRange(`B2:B${height+3}`);
	leftRange.setBorder(null, null, null, true, false, false);
	nextRange.setBorder(null, null, null, true, false, false);

	col = 4;
	for (const result in template) {
		let lett = LETTERS[col];
		let range = sheet.getRange(`${lett}2:${lett}${height+3}`);
		range.setBorder(null, null, null, true, false, false);
		col += 2;
	}

	let rightRange = sheet.getRange(`${LETTERS[col+2]}2:${LETTERS[col+2]}${height+1}`);
	rightRange.setBorder(null, null, null, true, false, false);
	let bottomRange = sheet.getRange(`${LETTERS[col-1]}${height+1}:${LETTERS[col+2]}${height+1}`);
	bottomRange.setBorder(null, null, true, null, false, false);

	let range = sheet.getRange(`B2:${LETTERS[col+2]}${height+1}`);
	let colors = createColorArr(col + 1, height);
	range.setBackgrounds(colors);


	col = 4;
	for (const result in template) {
		let column = sheet.getRange(`${LETTERS[col]}2:${LETTERS[col]}${height+2}`);
		column.setNumberFormat('00.0%');
		// range.setFontSize(11);
		col += 2;
	}
}

function addAverageRow(sheet, template, mainRangeValues) {
	let row = mainRangeValues.length + 1;
	let rangeValues = [
		['Averages'],
		['Totals']
	];

	rangeValues[0].push(`=AVERAGE(B2:B${row})`);
	rangeValues[1].push(`=SUM(B2:B${row})`);

	let col = 3;
	for (const result in template) {
		rangeValues[0].push(`=AVERAGE(${LETTERS[col]}2:${LETTERS[col]}${row})`);
		rangeValues[1].push(`=SUM(${LETTERS[col]}2:${LETTERS[col]}${row})`);
		rangeValues[0].push(`=AVERAGE(${LETTERS[col+1]}2:${LETTERS[col+1]}${row})`);
		rangeValues[1].push('');
		col += 2;
	}

	let range = sheet.getRange(`A${row+1}:${LETTERS[col-1]}${row+2}`);
	let range2 = sheet.getRange(`B${row+1}:${LETTERS[col-1]}${row+2}`);
	range.setValues(rangeValues);
	range.setNumberFormat('####');
	range.setFontWeight('bold');
	range.setFontSize(11);

	range2.setBackground(COLORS['darkGray']);
}

function fillInTemplate(personalData, template) {
	let templateCopy = {
		...template
	};
	personalData.forEach((val) => {
		let result = val[4];
		templateCopy[result]++;
	});
	return templateCopy;
}


function createResultsTemplate(data) {
	let template = {};
	for (const name in data) {
		let rows = data[name];
		rows.forEach((row) => {
			let result = row[4];
			template[result] = 0;
		});
	}
	return template;
}

function createIndividualSheetCallResults(callerID, personalData, ss) {
	ss.insertSheet(callerID);
	let sheet = ss.getSheetByName(callerID);

	let topRow = ['Voter ID', 'Voter Name', 'Voter Phone', 'Call Date', 'Result', 'Call Time', 'Time Diff'];
	let topRowRange = addRange(ss.getId(), topRow, 1, 1, callerID, 'row');

	let mainRange = addRange(ss.getId(), personalData, 2, 1, callerID);
	let colors = createColorArr(personalData[0].length, personalData.length);
	mainRange.setBackgrounds(colors);

	let timeInfo = addTimeDiffs(sheet, personalData, ss.getId(), callerID);
	formatIndividualSheet(sheet, personalData);
	return timeInfo;
}

function formatIndividualSheet(sheet, data) {
	sheet.setHiddenGridlines(true);

	sheet.setColumnWidth(1, 77);
	sheet.setColumnWidth(2, 131);
	sheet.setColumnWidth(3, 93);
	sheet.setColumnWidth(4, 71);
	sheet.setColumnWidth(5, 150);
	sheet.setColumnWidth(6, 92);
	sheet.setColumnWidth(7, 64);

	let topRow = sheet.getRange('A1:G1');
	topRow.setFontWeight('bold');
	topRow.setBackground(COLORS['darkGray']);

	let bottomRow = sheet.getRange(`A${data.length+1}:G${data.length+1}`);
	bottomRow.setBorder(null, null, true, null, false, false);

	let timeDiffRange = sheet.getRange(`G2:G${data.length+1}`);
	let timeDiffVals = timeDiffRange.getValues();

	colors = [];
	weights = [];

	c = 2;
	timeDiffVals.forEach((val) => {
		let num = Number(String(val).split(" ")[0]);
		if (num >= 25) {
			colors.push(['red']);
			weights.push(['bold']);
		} else if (num == 0) {
			colors.push(['yellow']);
			weights.push(['bold']);
		} else if (num > 2) {
			colors.push(['orange']);
			weights.push(['bold']);
		} else {
			if (c % 2 == 0) {
				colors.push(['white']);
				weights.push(['normal']);
			} else {
				colors.push([COLORS['gray']]);
				weights.push(['normal']);
			}
		}
		c++;
	});
	timeDiffRange.setBackgrounds(colors);
	timeDiffRange.setFontWeights(weights);
	timeDiffRange.setBorder(null, null, null, true, false, false);
}

function addTimeDiffs(sheet, data, ssID, callerID) {
	if (data.length == 1) {
		let range = sheet.getRange('F2');
		return [
			[0], range.getValues(), range.getValues()
		]
	}

	let timeRange = sheet.getRange(`F2:F${data.length+1}`);
	let timeVals = timeRange.getValues();

	let timeDiffsArr = [];
	let avg = 0;

	timeVals.forEach((cur, ind, arr) => {
		if (ind == arr.length - 1) {
			return;
		} else {
			let next = new Date('1970/01/01 ' + arr[ind + 1]);
			cur = new Date('1970/01/01 ' + cur);
			let timeDiff = next - cur;
			avg += timeDiff / 60000;
			timeDiffsArr.push([`${timeDiff/60000} mins`]);
		}
	});

	let timeDiffRange = addRange(ssID, timeDiffsArr, 2, 7, callerID, 'col');

	let startTime = timeVals[0];
	let endTime = timeVals[timeVals.length - 1];

	avg = Math.round((avg / timeDiffsArr.length + 1) * 100) / 100;
	let avgRow = ['Avg Time Diff', `${avg} mins`];
	let weights = ['bold', 'normal'];
	let avgRowRange = addRange(ssID, avgRow, 2, 8, callerID, 'row');
	avgRowRange.setFontWeights([weights]);

	return [timeDiffsArr, startTime, endTime];
}

function generateDataDict(arr) {
	let data = {};
	arr.forEach((val) => {
		let callerID = val[7];
		if (callerID == "" || callerID == "Caller Login") {
			return;
		} else {
			data[callerID] = [];
		}
	});
	console.log(`585`);
	console.log(arr[585]);
	console.log(arr.slice(580, 590));
	arr.forEach((val, ind) => {
		let voterID = val[0];
		let fullName = `${val[2]} ${val[3]}`;
		let phone = String(val[4]);
		phone = `(${phone.slice(0,3)}) ${phone.slice(3,6)}-${phone.slice(6,10)}`;
		let date = val[5];
		let time = val[6];
		let callerID = val[7];
		if (callerID == "" || callerID == "Caller Login") {
			return;
		} else {
			let result = val[8];
			if (callerID == undefined) {
				console.log(ind);
			}
			data[callerID].push([
				voterID, fullName, phone,
				date, result, time
			]);
		}
	});
	Logger.log(data);
	return sortDictByTime(data);
}


function generateTimeInfo(data) {

}


function sortDictByTime(dict) {
	// sorts individual data by time
	// in descending order
	for (const name in dict) {
		let data = dict[name];
		let sortedData = data.sort(function (a, b) {
			return new Date('1970/01/01 ' + a[5]) - new Date('1970/01/01 ' + b[5]);
		});
		dict[name] = sortedData;
	}
	return dict;
}