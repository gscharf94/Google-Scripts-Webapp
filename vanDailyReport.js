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
	Logger.log(`starting VAN daily.\nTeams:`);
	Logger.log(teams);
	
	let oss = SpreadsheetApp.openById(fileID);
	let sheets = oss.getSheets();

	Logger.log(`file id: ${fileID}`);

	let fileName = DriveApp.getFileById(fileID).getName();
	let dateStr = fileName.slice(0,6);
	dateStr = `${dateStr.slice(0,2)}-${dateStr.slice(2,4)}-${dateStr.slice(4,6)}`;

	let parentFolder = DriveApp.getFileById(fileID).getParents().next();
	let newFolder = parentFolder.createFolder(`${dateStr} GENERATED`);
	Logger.log(`generated new folder: ${newFolder.getName()}`);

	let nameList = [];
	for (const team in teams) {
		for (const name of teams[team]) {
			nameList.push(name);
		}
	}

	Logger.log(`namelist: ${nameList}`);	

	Logger.log(`Checking for errors..`);
	let output = checkSheetForErrors(teams, nameList, fileID);

	if (output != "") {
		Logger.log('Errors found');
		return output;
	}

	Logger.log(`Getting overview data`);
	let overviewData = extractOverviewData(fileID);
	Logger.log(overviewData);

	Logger.log(`getting time info`);
	let timeInfo = getTimeInfo(nameList, fileID);
	Logger.log(timeInfo);


	for (const name in overviewData) {
		overviewData[name]['timeInfo'] = timeInfo[name];
	}

	Logger.log('added timeinfo to overview data');

	let individualSheetData = {};

	for (const team in teams) {

		if (teams[team].length == 0) {
			continue;
		}

		Logger.log(`starting team: ${team}`);
		let teamSSID = createSpreadsheetNamed(newFolder, `${dateStr} TEAM ${team}`);
		createOverviewPage(overviewData, teams[team], teamSSID, params);

		var teamIndividualData = {};

		for (const name of teams[team]) {
			Logger.log(`Creating individual sheet for: ${name}`);
			let individualData = extractIndividualData(fileID, name);
			teamIndividualData[name] = individualData;
			createIndividualSheet(teamSSID, individualData, name, timeInfo[name]['timeDiffs'], params);
		}

		addNameHyperlinks(teams[team], teamSSID);
		individualSheetData[team] = teamIndividualData;
	}

	Logger.log(`Starting visuals`);
	createGraphs(teams, overviewData, null, newFolder, dateStr);
	Logger.log('Creating team comparison sheet');
	createTeamComparisonSheet(teams, overviewData, newFolder, dateStr);
}

function addNameHyperlinks(nameList, ssid) {
	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Overview');

	let nameRange = sheet.getRange(`A3:A${2+nameList.length}`);
	
	function getHyperLink(name, ssid) {
		Logger.log(`name: ${name}`);
		let ss = SpreadsheetApp.openById(ssid);
		let sheet = ss.getSheetByName(name);
		let link = `https://docs.google.com/spreadsheets/d/${ssid}/edit#gid=${sheet.getSheetId()}`;
		return `=HYPERLINK("${link}","${name}")`;
	}

	let rangeData = [];
	Logger.log(`range data: ${nameRange.getValues()}`);
	for (const name of nameRange.getValues()) {
		rangeData.push([getHyperLink(name[0], ssid)]);
	}

	nameRange.setFormulas(rangeData);
}

function createTeamComparisonSheet(teams, data, parentFolder, dateStr) {
	let ssid = createSpreadsheetNamed(parentFolder, `${dateStr} Team Comparison`);

	let aggregateData = {};
	for (const team in teams) {
		if (teams[team].length == 0) {
			continue;
		}

		let teamData = {
			totalCalls: 0,
			canvassed: 0,
			leftMessage: 0,
			refused: 0,
			otherLanguage: 0,
			notHome: 0,
			disconnected: 0,
			moved: 0,
			other: 0,
		}

		for (const name of teams[team]) {
			teamData['totalCalls'] += Number(data[name]['totalCalls']);
			teamData['canvassed'] += Number(data[name]['canvassed']);
			teamData['leftMessage'] += Number(data[name]['leftMessage']);
			teamData['refused'] += Number(data[name]['refused']);
			teamData['otherLanguage'] += Number(data[name]['otherLanguage']);
			teamData['notHome'] += Number(data[name]['notHome']);
			teamData['disconnected'] += Number(data[name]['disconnected']);
			teamData['moved'] += Number(data[name]['moved']);
			teamData['other'] += Number(data[name]['other']);
		}

		aggregateData[team] = teamData;
	}

	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Overview');
	sheet.setHiddenGridlines(true);

	let rangeData = [[
		'Team',
		'Total Calls',
		"Canv'd",
		'Left Mess',
		'Ref',
		'Other Lang',
		'Not Home',
		'Disc',
		'Mvd',
		'Other'
	]];

	for (const team in aggregateData) {
		rangeData.push([
			team,
			aggregateData[team].totalCalls,
			aggregateData[team].canvassed,
			aggregateData[team].leftMessage,
			aggregateData[team].refused,
			aggregateData[team].otherLanguage,
			aggregateData[team].notHome,
			aggregateData[team].disconnected,
			aggregateData[team].moved,
			aggregateData[team].other,
		]);
	}

	let mainRange = addRange(ssid, rangeData, 2, 2, 'Overview');
	mainRange.setBorder(true, true, true, true, false, false);

	let topRowRange = sheet.getRange('B2:K2')
	topRowRange
		.setBackground(COLORS['middleGray'])
		.setHorizontalAlignment('center')
		.setFontWeight('bold')
		.setBorder(true, true, true, true, false, false)
		.setWrap(true);

	sheet.setRowHeight(2, 41);
	sheet.setColumnWidth(1, 40);
	sheet.setColumnWidth(2, 100);
	sheet.setColumnWidth(3, 60);
	sheet.setColumnWidth(4, 60);
	sheet.setColumnWidth(5, 60);
	sheet.setColumnWidth(6, 60);
	sheet.setColumnWidth(7, 60);
	sheet.setColumnWidth(8, 60);
	sheet.setColumnWidth(9, 60);
	sheet.setColumnWidth(10, 60);
	sheet.setColumnWidth(11, 60);

	let c = 0;
	for (const team in aggregateData) {
		c++;
	}

	let leftRange = sheet.getRange(`A1:A${c+1}`);
	leftRange
		.setHorizontalAlignment('right');

	let chartRange = sheet.getRange('B2:C15');

	function createPieChart(sheet, range) {
		let chart = sheet.newChart()
			.setChartType(Charts.ChartType.PIE)
			.setOption('width', 500)
			.setOption('height', 500)
			.setNumHeaders(1)
			.addRange(range)
			.setOption('legend', { position: "left" })
			.setOption('pieSliceText', 'value-and-percentage')
			.setOption('title','Total Calls per Team')
			.setPosition(8, 1, 0, 0);

		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart
	}

	createPieChart(sheet, chartRange);

	let chart2Range = sheet.getRange(`B2:B15`);
	let chart2Range2 = sheet.getRange('D2:D15');

	function createPieChart2(sheet, r1, r2) {
		let chart = sheet.newChart()
			.setChartType(Charts.ChartType.PIE)
			.setOption('width', 500)
			.setOption('height', 500)
			.setNumHeaders(1)
			.addRange(r1)
			.addRange(r2)
			.setOption('legend', { position: "left" })
			.setOption('pieSliceText', 'value-and-percentage')
			.setOption('title','Total Canvassed per Team')
			.setPosition(8, 9, 0, 0);

		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart
	}

	createPieChart2(sheet, chart2Range, chart2Range2);
}


function createGraphs(teams, overviewData, teamIndividualData, parentFolder, dateStr) {
	
	let visualsSSID = createSpreadsheetNamed(parentFolder, `${dateStr} Visuals`);
	createHourlyCanvStats(visualsSSID, overviewData, teams);
	createTimeDiffsGraph(visualsSSID, overviewData, teams);
	createHoursWorkedGraph(visualsSSID, overviewData, teams);
}

function createTimeDiffsGraph(ssid, data, teams) {
	let ss = SpreadsheetApp.openById(ssid);
	ss.insertSheet('Time Diffs');
	let sheet = ss.getSheetByName('Time Diffs');
	sheet.setHiddenGridlines(true);

	let minorThreshold = 5;
	let majorThreshold = 25;

	let rangeData = [[
		'Name',
		`+${5}mins Time Diffs`,
	]];

	for (const team in teams) {
		for (const name of teams[team]) {
			let diffs = data[name]['timeInfo']['timeDiffs'].filter( (val) => val > minorThreshold ).length;
			rangeData.push([
				name,
				diffs,
			]);
		}
	}

	let range = addRange(ssid, rangeData, 1, 1, 'Time Diffs');

	function createTimeDiffsChart(sheet, range) {
		let chart = sheet.newChart();

		chart
			.addRange(range)
			.setChartType(Charts.ChartType.BAR)
			.setOption('width', 1200)
			.setOption('height', 1200)
			.setOption('legend', { position: "none" })
			.setOption('titlePosition', 'none')
			.setNumHeaders(1)
			.setOption('series.0.dataLabel', 'value')
			.setOption('annotations.alwaysOutside', true)
			.setOption('series.0.annotations.textStyle.fontSize', 15)
			.setOption('chartArea', {top: "2.5%", height: "97.5%"})
			.setPosition(1,1,0,0);

			let i = 0;
			let c = 0;
			let colors = ['brown','orange','teal','green','olive','crimson','red'];
			for (const team in teams) {
				for (const name of teams[team]) {
					chart.setOption(`series.0.items.${i}.color`, colors[c]);
					i++;
				}
				c++;
			}

			let finishedChart = chart.build();
			sheet.insertChart(finishedChart);
			return finishedChart;
	}

	let chart = createTimeDiffsChart(sheet, range);
}

function createHoursWorkedGraph(ssid, data, teams) {
	let ss = SpreadsheetApp.openById(ssid);
	ss.insertSheet('Hours Worked');
	let sheet = ss.getSheetByName('Hours Worked');
	sheet.setHiddenGridlines(true);

	let rangeData = [[
		'Name',
		'Hours Worked',
	]];

	for (const team in teams) {
		for (const name of teams[team]) {
			let { sTimeStr, eTimeStr, hrsWorked, hrlyAttempts, hrlyCanv } = getHourlyAvgData(data, name);
			rangeData.push([
				name,
				hrsWorked,
			]);
		}
	}

	let range = addRange(ssid, rangeData, 1, 1, 'Hours Worked');
	range.setNumberFormat('0.0');

	function createChart(sheet, range) {
		let chart = sheet.newChart();
		chart
			.addRange(range)
			.setChartType(Charts.ChartType.BAR)
			.setOption('width', 1200)
			.setOption('height', 1200)
			.setOption('legend', { position: "none" })
			.setOption('titlePosition', 'none')
			.setNumHeaders(1)
			.setOption('series.0.dataLabel', 'value')
			.setOption('annotations.alwaysOutside', true)
			.setOption('series.0.annotations.textStyle.fontSize', 15)
			.setOption('chartArea', {top: "2.5%", height: "97.5%"})
			.setPosition(1,1,0,0);

		let i = 0;
		let c = 0;
		let colors = ['brown','orange','teal','green','olive','crimson','red'];
		for (const team in teams) {
			for (const name of teams[team]) {
				chart.setOption(`series.0.items.${i}.color`, colors[c]);
				i++;
			}
			c++;
		}

		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart;
	}
	let chart = createChart(sheet, range);
}

function createHourlyCanvStats(ssid, data, teams) {
	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Hourly Calls + Canv');
	sheet.setHiddenGridlines(true);

	let rangeData = [[
		'Name',
		'Canv',
		'Total',
	]];

	for (const team in teams) {
		for (const name of teams[team]) {
			let { sTimeStr, eTimeStr, hrsWorked, hrlyAttempts, hrlyCanv } = getHourlyAvgData(data, name);
			rangeData.push([
				name,
				hrlyCanv,
				hrlyAttempts,
			]);
		}
	}

	let range = addRange(ssid, rangeData, 1, 1, 'Hourly Calls + Canv');
	range.setNumberFormat('0.0');

	function createAvgHourlyGraph(sheet, range) {
		let chart = sheet.newChart();
		chart
			.addRange(range)
			.setChartType(Charts.ChartType.BAR)
			.setOption('width', 1200)
			.setOption('height', 1200)
			.setOption('legend', { position: "none" })
			.setOption('titlePosition', 'none')
			.setNumHeaders(1)
			.setOption('isStacked', true)
			.setOption('series.0.dataLabel', 'value')
			.setOption('series.1.dataLabel', 'value')
			.setOption('bar.groupWidth', "50")
			.setOption('series.0.annotations.textStyle.fontSize', 15)
			.setOption('series.0.annotations.highContract', false)
			.setOption('series.1.annotations.textStyle.fontSize', 15)
			.setOption('annotations.alwaysOutside', true)
			// .setOption('chartArea', {left: "12%", top: "2.5%", width: "70%", height: "97.5%"})
			.setOption('chartArea', {top: "2.5%", height: "97.5%"})
			.setPosition(1,1,0,0);

		let c = 0;
		let i = 0;
		let canvColorArr = ['brown','orange','teal','green','olive','crimson','red'];
		let totalColorArr = ['lightcoral','moccasin','turquoise','limegreen','yellowgreen','palevioletred','salmon'];
		for (const team in teams) {
			for (const name of teams[team]) {
				chart.setOption(`series.0.items.${i}.color`, canvColorArr[c]);
				chart.setOption(`series.1.items.${i}.color`, totalColorArr[c]);
				i++;
			}
			c++;
		}
		
		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart;
	}

	let chart = createAvgHourlyGraph(sheet, range);

	let topCanv = rangeData.slice(1,);
	let topCalls = rangeData.slice(1,);

	topCanv.sort( (a,b) => {
		return b[1] - a[1];
	});

	topCalls.sort( (a,b) => {
		return b[2] - a[2];
	});

	let topCanvData = [['Name','Canv']];
	for (let i = 0; i < 5; i++) {
		topCanvData.push([
			topCanv[i][0],
			topCanv[i][1],
		]);
	}

	let topCallsData = [['Name', 'Calls']];
	for (let i = 0; i < 5; i++) {
		topCallsData.push([
			topCalls[i][0],
			topCalls[i][2],
		]);
	}

	let topCanvRange = addRange(ssid, topCanvData, 3, 13, 'Hourly Calls + Canv');
	topCanvRange.setNumberFormat('0.0');
	topCanvRange.setBorder(true, true, true, true, false, false);
	
	let topRowCanv = sheet.getRange('M3:N3');
	topRowCanv
		.setBackground(COLORS['middleGray'])
		.setHorizontalAlignment('center')
		.setFontWeight('bold')
		.setBorder(true, true, true, true, false, false);

	sheet.getRange('M2')
		.setValue('Top 5 Hourly Canvassed')
		.setBackground(COLORS['middleGray'])
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setBorder(true, true, true, true, false, false);

	sheet.getRange('M2:N2').merge();
	
	let topCallsRange = addRange(ssid, topCallsData, 11, 13, 'Hourly Calls + Canv');
	topCallsRange.setNumberFormat('0.0');
	topCallsRange.setBorder(true, true, true, true, false, false);

	let topRowCall = sheet.getRange('M11:N11');
	topRowCall
		.setBackground(COLORS['middleGray'])
		.setHorizontalAlignment('center')
		.setFontWeight('bold')
		.setBorder(true, true, true, true, false, false);

	sheet.setColumnWidth(13, 150)
	sheet.setColumnWidth(14, 65)

	sheet.getRange('M10')
		.setValue('Top 5 Hourly Calls')
		.setBackground(COLORS['middleGray'])
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setBorder(true, true, true, true, false, false);

	sheet.getRange('M10:N10').merge()
}

function extractIndividualData(ssid, name) {
	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName(name);

	let addressCol = sheet.getRange("A1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let personCol = sheet.getRange("B1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let timeCol = sheet.getRange("C1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();

	let nhCol = sheet.getRange(`D1:D${addressCol.length}`).getValues();
	let refCol = sheet.getRange(`E1:E${addressCol.length}`).getValues();
	let mvdCol = sheet.getRange(`F1:F${addressCol.length}`).getValues();
	let decCol = sheet.getRange(`G1:G${addressCol.length}`).getValues();
	let canvCol = sheet.getRange(`H1:H${addressCol.length}`).getValues();

	let arr = [];

	for (let i = 1; i < addressCol.length; i++) {
		let row = [
			addressCol[i],
			personCol[i],
			timeCol[i],
			nhCol[i],
			refCol[i],
			mvdCol[i],
			decCol[i],
			canvCol[i],
		];
		arr.push(row);
	}

	return arr;
}

function createIndividualSheet(ssid, data, name, tDiffs, params) {
	let ss = SpreadsheetApp.openById(ssid);
	ss.insertSheet(name);
	let sheet = ss.getSheetByName(name);
	sheet.setHiddenGridlines(true);

	let minorThreshold = params[0];
	let majorThreshold = params[1];


	let toWrite = [[
		'Address',
		'Name',
		'Time',
		'Diff',
		'NH',
		'Ref',
		'Mvd',
		'Dec',
		'Canv',
	]];

	tDiffs.unshift('');

	data.forEach( (row, ind, arr) => {
		let dateStr = new Date(row[2]);
		dateStr = `${String(dateStr.getHours()).padStart(2,'0')}:${String(dateStr.getMinutes()).padStart(2,'0')}`;
		let tDiffStr = `${tDiffs[ind]} mins`;
		if (ind == 0) {
			tDiffStr = '';
		}
		
		let appendRow = [
			row[0],
			row[1],
			dateStr,
			tDiffStr,
			row[3],
			row[4],
			row[5],
			row[6],
			row[7],
		];

		toWrite.push(appendRow);
	});

	let mainRange = addRange(ssid, toWrite, 1, 1, name);

	let colors = createColorArr(9, toWrite.length-1);
	colors.unshift([
		COLORS['middleGray'], COLORS['middleGray'],
		COLORS['middleGray'], COLORS['middleGray'],
		COLORS['middleGray'], COLORS['middleGray'],
		COLORS['middleGray'], COLORS['middleGray'],
		COLORS['middleGray'],
	]);

	mainRange.setBackgrounds(colors);

	sheet.getRange('A1:I1')
		.setFontWeight('bold')
		.setHorizontalAlignment('center');

	sheet.getRange(`E2:I${toWrite.length}`)
		.setHorizontalAlignment('center');


	sheet.setColumnWidth(1, 190);
	sheet.setColumnWidth(2, 160);
	sheet.setColumnWidth(3, 50);
	sheet.setColumnWidth(4, 55);
	sheet.setColumnWidth(5, 40);
	sheet.setColumnWidth(6, 40);
	sheet.setColumnWidth(7, 40);
	sheet.setColumnWidth(8, 40);
	sheet.setColumnWidth(9, 40);

	let timeCol = sheet.getRange("D3").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();



	let c = 0;
	let timeColors = [];
	for (let i = 0; i < timeCol.length; i++) {
		c++;
		let val = Number(`${timeCol[i]}`.split(" ")[0]);
		if (val >= majorThreshold) {
			timeColors.push([COLORS['red']]);
		} else if (val >= minorThreshold) {
			timeColors.push([COLORS['orange']]);
		} else if (val == 0) {
			timeColors.push([COLORS['yellow']]);
		} else {
			if (i%2 == 0) {
				timeColors.push([COLORS['gray']]);
			} else {
				timeColors.push(['white']);
			}
		}
	}

	let timeDiffRange = sheet.getRange(`D3:D${c+2}`)
	timeDiffRange.setBackgrounds(timeColors);


	// adding hyperlinks back to overview pages
	let topLinkRange = sheet.getRange('J1');
	let link = `https://docs.google.com/spreadsheets/d/${ssid}/edit#gid=0`
	topLinkRange.setFormula(`=HYPERLINK("${link}", "Back to Overview")`);
	
	let bottomLinkRange = sheet.getRange(`A${timeColors.length+1}`);
	bottomLinkRange.setFormula(`=HYPERLINK("${link}", "Back to Overview")`);
}


function createOverviewPage(data, team, ssid, params) {
	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Overview');

	let minorThreshold = params[0];
	let majorThreshold = params[1];

	sheet.getRange('C1').setValue('Canvassed');
	sheet.getRange('C1:D1').merge();
	
	sheet.getRange('E1').setValue('Not Home');
	sheet.getRange('E1:F1').merge();
	
	sheet.getRange('G1').setValue('Left Message');
	sheet.getRange('G1:H1').merge();

	sheet.getRange('I1').setValue('Refused');
	sheet.getRange('I1:J1').merge();
	
	sheet.getRange('K1').setValue('Other Language');
	sheet.getRange('K1:L1').merge();
	
	sheet.getRange('M1').setValue('Disconnect');
	sheet.getRange('M1:N1').merge();
	
	sheet.getRange('O1').setValue('Moved');
	sheet.getRange('O1:P1').merge();
	
	sheet.getRange('Q1').setValue('Other');
	sheet.getRange('Q1:R1').merge();
	
	let dataToWrite = [];
	dataToWrite.push([
		'Name',
		'Total Calls',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
	]);

	function createNameHyperlink(name, ssid, sheetId) {

	}


	for (const name of team) {
		let row = [
			name,
			data[name]['totalCalls'],
			data[name]['canvassed'],
			Number(data[name]['canvassed']) / Number(data[name]['totalCalls']), 
			data[name]['notHome'],
			Number(data[name]['notHome']) / Number(data[name]['totalCalls']), 
			data[name]['leftMessage'],
			Number(data[name]['leftMessage']) / Number(data[name]['totalCalls']), 
			data[name]['refused'],
			Number(data[name]['refused']) / Number(data[name]['totalCalls']), 
			data[name]['otherLanguage'],
			Number(data[name]['otherLanguage']) / Number(data[name]['totalCalls']), 
			data[name]['disconnected'],
			Number(data[name]['disconnected']) / Number(data[name]['totalCalls']), 
			data[name]['moved'],
			Number(data[name]['moved']) / Number(data[name]['totalCalls']), 
			data[name]['other'],
			Number(data[name]['other']) / Number(data[name]['totalCalls']), 
		];
		dataToWrite.push(row);
	}

	let topRange = addRange(ssid, dataToWrite, 2, 1, 'Overview');

	let row = 3 + team.length;
	let sumTotalRows = [['Team Averages'],['Team Totals']];
	let sumTotalRange = addRange(ssid, sumTotalRows, row, 1, 'Overview', 'col');

	let avgRow = [];
	let sumRow = [];
	for (let c = 2; c < 19; c++) {
		let col = LETTERS[c];
		let range = `${col}3:${col}${row-1}`;
		avgRow.push(`=AVERAGE(${range})`);
		if (c > 3 && c%2 == 0) {
			sumRow.push('');
		} else {
			sumRow.push(`=SUM(${range})`);
		}
	}
	
	let sumTotalDataRange = sheet.getRange(`B${row}:R${row+1}`);
	sumTotalDataRange.setFormulas([avgRow, sumRow]);

	row = row + 2;

	sheet.getRange(`C${row}`).setValue(`Canvassed`);
	sheet.getRange(`C${row}:D${row}`).merge();
	
	sheet.getRange(`E${row}`).setValue(`Not Home`);
	sheet.getRange(`E${row}:F${row}`).merge();
	
	sheet.getRange(`G${row}`).setValue(`Left Message`);
	sheet.getRange(`G${row}:H${row}`).merge();

	sheet.getRange(`I${row}`).setValue(`Refused`);
	sheet.getRange(`I${row}:J${row}`).merge();
	
	sheet.getRange(`K${row}`).setValue(`Other Language`);
	sheet.getRange(`K${row}:L${row}`).merge();
	
	sheet.getRange(`M${row}`).setValue(`Disconnected`);
	sheet.getRange(`M${row}:N${row}`).merge();
	
	sheet.getRange(`O${row}`).setValue(`Moved`);
	sheet.getRange(`O${row}:P${row}`).merge();
	
	sheet.getRange(`Q${row}`).setValue(`Other`);
	sheet.getRange(`Q${row}:R${row}`).merge();


	let totalsArr = [];
	totalsArr.push([
		'', 'Total Calls',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
		'Total', 'Average',
	]);

	totalsArr.push([
		'Program Totals',
		data.totals.totalCalls,
		data.totals.canvassed,
		Number(data.totals.canvassed) / Number(data.totals.totalCalls),
		data.totals.notHome,
		Number(data.totals.notHome) / Number(data.totals.totalCalls),
		data.totals.leftMessage,
		Number(data.totals.leftMessage) / Number(data.totals.totalCalls),
		data.totals.refused,
		Number(data.totals.refused) / Number(data.totals.totalCalls),
		data.totals.otherLanguage,
		Number(data.totals.otherLanguage) / Number(data.totals.totalCalls),
		data.totals.disconnected,
		Number(data.totals.disconnected) / Number(data.totals.totalCalls),
		data.totals.moved,
		Number(data.totals.moved) / Number(data.totals.totalCalls),
		data.totals.other,
		Number(data.totals.other) / Number(data.totals.totalCalls),
	]);

	let programTotalRange = addRange(ssid, totalsArr, row + 1, 1, 'Overview');
	
	row = row + 4;

	let secondDataArr = [[
		'', 'Start Time', 'End Time',
		'Hrs Worked', 'Attmpts/Hr',
		'Canv/Hr', `${minorThreshold}+ Diffs`,
		'Avg Time',
	]];

	for (const name of team) {
		let { sTimeStr, eTimeStr, hrsWorked, hrlyAttempts, hrlyCanv } = getHourlyAvgData(data, name);

		let minorTimeDiffs = data[name]['timeInfo']['timeDiffs'].filter( (val) => val > minorThreshold ).length;

		secondDataArr.push([
			name,
			sTimeStr,
			eTimeStr,
			hrsWorked,
			hrlyAttempts,
			hrlyCanv,
			minorTimeDiffs,
			data[name]['timeInfo']['avgTimeDiff'],
		]);
	}

	let secondDataRange = addRange(ssid, secondDataArr, row, 1, 'Overview');

	row = row + secondDataArr.length + 1;

	let thirdDataArr = [];

	sheet.getRange(`B${row}`).setValue(`${majorThreshold}+ Min Breaks`);
	sheet.getRange(`B${row}:D${row}`).merge();
	sheet.getRange(`B${row}`).setBorder(true, true, true, true, false, false);


	for (const name of team) {
		let majorBreaks = data[name]['timeInfo']['timeDiffs'].filter( (val) => val > majorThreshold );
		let sum = majorBreaks.reduce( (a,b) => a + b, 0);
		let row = [name, 'Total'];
		if (sum == 0) {
			row.push('0 mins');
		} else {
			row.push(`${sum} mins`);
		}

		if (majorBreaks.length == 0) {
			row.push('0 breaks');
		} else {
			row.push(`${majorBreaks.length} breaks`);
		}

		thirdDataArr.push(row);
	}

	let thirdDataRange = addRange(ssid, thirdDataArr, row + 1, 1, 'Overview');
	
	formatOverviewSheet(sheet, sumTotalDataRange, programTotalRange, secondDataRange, thirdDataRange, params, team.length);
}

function getHourlyAvgData(data, name) {
	let sTime = data[name]['timeInfo']['startTime'];
	let sTimeStr = `${String(sTime.getHours()).padStart(2, '0')}:${String(sTime.getMinutes()).padStart(2, '0')}`;
	let eTime = data[name]['timeInfo']['endTime'];
	let eTimeStr = `${String(eTime.getHours()).padStart(2, '0')}:${String(eTime.getMinutes()).padStart(2, '0')}`;
	let hrsWorked = eTime - sTime;
	hrsWorked = (hrsWorked / 60000) / 60;

	let hrlyAttempts = data[name]['totalCalls'] / hrsWorked;
	let hrlyCanv = data[name]['canvassed'] / hrsWorked;
	return { sTimeStr, eTimeStr, hrsWorked, hrlyAttempts, hrlyCanv };
}

function formatOverviewSheet(sheet, sumTotalRange, programTotalRange, secondRange, thirdRange, params, n) {

	sheet.setHiddenGridlines(true);

	sheet.getRange('C1:R1')
		.setHorizontalAlignment('center')
		.setFontWeight('bold');

	sheet.getRange(`C${n+5}:R${n+5}`)
		.setHorizontalAlignment('center')
		.setFontWeight('bold');
		
	sheet.getRange('B2:R2')
	.setHorizontalAlignment('center')
	.setFontWeight('bold')
	.setBackground(COLORS['middleGray'])
	.setBorder(true, true, true, true, false, false);
	
	sheet.getRange(`B${n+6}:R${n+6}`)
	.setHorizontalAlignment('center')
	.setBackground(COLORS['middleGray'])
	.setFontWeight('bold')
	.setBorder(true, true, true, true, false, false);
	
	let firstColRange = sheet.getRange(`A2:A${n*10}`);
	let fontWeights = [];
	for (const val of firstColRange.getValues()) {
		if (val == "Team Averages" || val == "Team Totals" || val == "Program Totals" || val == "Name") {
			fontWeights.push( ['bold'] );
		} else {
			fontWeights.push( ['normal'] );
		}
	}

	firstColRange
		.setHorizontalAlignment('right')
		.setFontWeights(fontWeights);

	let overviewRange = sheet.getRange(`B3:R${n+2}`);
	let colors = createColorArr(17, n);
	overviewRange
		.setBackgrounds(colors)
		.setHorizontalAlignment('center')
		.setBorder(true, true, true, true, false, false);

	let cRow = n+3;

	let avgRange = sheet.getRange(`B${cRow}:R${cRow}`);
	
	numFormats = [[
		'#.#',
		'#.#', '0.0%',
		'#.#', '0.0%',
		'#.#', '0.0%',
		'#.#', '0.0%',
		'#.#', '0.0%',
		'#.#', '0.0%',
		'#.#', '0.0%',
		'#.#', '0.0%',
	]];
	
	avgRange.setNumberFormats(numFormats);


	let sumTotalsRange = sheet.getRange(`B${cRow}:R${cRow+1}`);
	sumTotalsRange
		.setFontWeight('bold')
		.setBackground(COLORS['darkGray'])
		.setHorizontalAlignment('center')
		.setFontSize(11)
		.setFontColor('white')
		.setBorder(true, true, true, true, false, false);

	cRow += 2;

	programTotalRange = sheet.getRange(`B${cRow+2}:R${cRow+2}`);
	programTotalRange
		.setBorder(true, true, true, true, false, false)
		.setHorizontalAlignment('center');

	cRow += 3;

	let topRowRange = sheet.getRange(`B${cRow+1}:H${cRow+1}`);
	topRowRange
		.setBorder(true, true, true, true, false, false)
		.setBackground(COLORS['middleGray'])
		.setHorizontalAlignment('center');

	sheet.getRange(`B${cRow+2}:H${cRow+2}`)
		.setHorizontalAlignment('center')
		.setBorder(true, true, true, true, false, false);

	let secondDataRange = sheet.getRange(`B${cRow+1}:H${cRow+1+n}`);
	secondDataRange
		.setBorder(true, true, true, true, false, false)
		.setHorizontalAlignment('center')
		// .setFontWeight('bold');

	cRow = cRow + 1 + n + 2;

	sheet.getRange(`B${cRow}`)
		.setFontWeight('bold')
		.setBackground(COLORS['middleGray'])
		.setHorizontalAlignment('center');

	let lastRange = sheet.getRange(`B${cRow+1}:D${cRow+n}`);
	lastRange
		.setHorizontalAlignment('center')
		.setBorder(true, true, true, true, false, false);
	

	let z = 3;

	let cols = ['D', 'F', 'H', 'J', 'L', 'N', 'P', 'R'];
	for (const col of cols) {
		sheet.getRange(`${col}${z}:${col}${z+n-1}`)
			.setNumberFormat('0.0%');
		sheet.getRange(`${col}${n+7}`).setNumberFormat('0.0%');
	}

	z = n + 7;

	let stupidRange = sheet.getRange(`B${n+9}:H${n+9}`);
	stupidRange
		.setFontWeight('bold')
		.setBorder(true, true, true, true, false, false);
	
	sheet.getRange(`D${z+3}:D${z+2+n}`).setNumberFormat('#.#');
	sheet.getRange(`E${z+3}:E${z+2+n}`).setNumberFormat('#.#');
	sheet.getRange(`F${z+3}:F${z+2+n}`).setNumberFormat('#.#');
	sheet.getRange(`H${z+3}:H${z+2+n}`).setNumberFormat('#.#');

	sheet.setColumnWidth(1, 153);

	for (let i = 3; i < 20; i++) {
		sheet.setColumnWidth(i, 75);
	}
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

	let startTime = new Date(timeCol[0]);
	let endTime = new Date(timeCol[timeCol.length-1]);
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

	outputData.totals = {
		totalCalls: 0,
		canvassed: 0,
		leftMessage: 0,
		refused: 0,
		otherLanguage: 0,
		notHome: 0,
		disconnected: 0,
		moved: 0,
		other: 0,
	}

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

		outputData.totals.totalCalls += Number(totalCol[i]);
		outputData.totals.canvassed += Number(canvCol[i]);
		outputData.totals.leftMessage += Number(leftMessageCol[i]);
		outputData.totals.refused += Number(refusedCol[i]);
		outputData.totals.otherLanguage += Number(otherLanguageCol[i]);
		outputData.totals.notHome += Number(notHomeCol[i]);
		outputData.totals.disconnected += Number(disconnectedCol[i]);
		outputData.totals.moved += Number(movedCol[i]);
		outputData.totals.other += Number(otherCol[i]);
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