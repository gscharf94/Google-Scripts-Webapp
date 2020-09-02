function callerDetailsStart(fileID, fileName) {
	let file = DriveApp.getFileById(fileID);
	let parentFolder = file.getParents().next();
	let blob = file.getBlob();
	let rawText = blob.getDataAsString('utf8');

	let parsedCSV = parseCSV(rawText);
	let data = createDict(parsedCSV);
	let sheetID = createSpreadsheet(fileName, parentFolder,'CD-');
	let sortedData = createToWriteArr(data);

	let overviewSheet = setupOverviewSheet(sheetID, sortedData);
	let firstChart = createFirstChart(overviewSheet);
	let secondChart = createSecondChart(overviewSheet);

	let tdOutput = createTimeDiffSheet(sheetID, data);
	let tdSheet = tdOutput[0];
	let tdRange = tdOutput[1];
	let timeFlagChart = createTimeFlagChart(tdSheet, tdRange);

	
	return SpreadsheetApp.openById(sheetID).getUrl();

	
}

function createTimeFlagChart(sheet, range) {
	let chart = sheet.newChart();
	chart
		.addRange(range)
		.setChartType(Charts.ChartType.BAR)
		.setOption('width', 1000)
		.setOption('height', 1000)
		.setOption('legend', { position: "none" })
		.setOption('titlePosition', 'none')
		.setPosition(1,4,0,0);
	
	let stackedChart = chart.asBarChart().setStacked();
	let finishedChart = stackedChart.build();
	sheet.insertChart(finishedChart);
	return finishedChart;
}

function createTimeDiffSheet(sheetID, data) {
	let ss = SpreadsheetApp.openById(sheetID);
	ss.insertSheet('Time Flags');
	let sheet = ss.getSheetByName('Time Flags');
	sheet.setHiddenGridlines(true);

	function compileData() {
		let arr = [];
		for (const name in data) {
			let row = [
				name,
				Number(data[name]['inWrap']),
				Number(data[name]['inNotReady']),
			];
			arr.push(row);
		}

		let sortedData = arr.sort( (a,b) =>
			a[0][0].toLowerCase().charCodeAt() - b[0][0].toLowerCase().charCodeAt()
		);

		return sortedData;
	}

	let sortedData = compileData();
	sortedData.unshift(['Name','Wrap Up','Not Ready']);

	let range = addRange(sheetID, sortedData, 1, 1, 'Time Flags');
	range.setHorizontalAlignment('center');
	range.setFontWeight('bold');

	return [sheet, range];
}

function createDict(arr) {
	let dict = {};
	arr.forEach( (row, ind) => {
		if(ind == 0) {
			return;
		}
		dict[row[2]] = {
			'login':row[1],
			'email':row[3],
			'inCall':row[5],
			'inWrap':row[6],
			'inReady':row[7],
			'inNotReady':row[8],
			'totalCalls':row[9],
		  };
	});
	return dict;
}

function createFirstChart(sheet) {
	let chartBuilder = sheet.newChart();
	let dataRange = sheet.getRange('J2:M7');

	chartBuilder
		.addRange(dataRange)
		.setChartType(Charts.ChartType.BAR)
		.setPosition(8,9,0,0);

	let chart = chartBuilder.build();
	sheet.insertChart(chart);
	return chart;
}

function createSecondChart(sheet) {
	let chartBuilder = sheet.newChart();
	let dataRange = sheet.getRange('J29:J33');

	chartBuilder
		.addRange(dataRange)
		.setChartType(Charts.ChartType.COLUMN)
		.setPosition(34,9,0,0);

	let chart = chartBuilder.build();
	sheet.insertChart(chart);
	return chart;
}

function setupOverviewSheet(sheetID, sortedData) {
	// setups overview sheet
	// returns sheet obj
	let ss = SpreadsheetApp.openById(sheetID);
	let sheet = ss.getSheets()[0];
	sheet.setHiddenGridlines(true);
	sheet.setName('Overview');

	let topRow = ['Login','Name','Email','Wrap Up','Not Ready','Call','Ready','Total'];

	let topRange = addRange(sheetID, topRow, 1, 1, sheet.getName(), 'row');
	topRange.setFontWeight('bold');
	topRange.setFontSize(11);
	topRange.setBackground(COLORS['darkGray']);
	topRange.setHorizontalAlignment('center');

	let mainRange = addRange(sheetID, sortedData, 2, 1, sheet.getName());
	let colors = createColorArr(8, sortedData.length);
	mainRange.setBackgrounds(colors);

	sheet.setColumnWidth(1,108);
	sheet.setColumnWidth(2,177);
	sheet.setColumnWidth(3,177);
	sheet.setColumnWidth(4,74);
	sheet.setColumnWidth(5,74);
	sheet.setColumnWidth(6,74);
	sheet.setColumnWidth(7,74);
	sheet.setColumnWidth(8,74);
	sheet.setColumnWidth(10,200);
	sheet.setColumnWidth(11,70);
	sheet.setColumnWidth(12,70);
	sheet.setColumnWidth(14,200);
	
	
	let avgRow = [
		'',
		'',
		'AVERAGE',
		`=AVERAGE(D2:D${sortedData.length})`,
		`=AVERAGE(E2:E${sortedData.length})`,
		`=AVERAGE(F2:F${sortedData.length})`,
		`=AVERAGE(G2:G${sortedData.length})`,
		`=AVERAGE(H2:H${sortedData.length})`,
	];
	let avgRowRange = addRange(sheetID, avgRow, sortedData.length+1, 1, sheet.getName(), 'row');
	avgRowRange.setFontSize(11);
	avgRowRange.setFontWeight('bold');
	avgRowRange.setBackground(COLORS['darkGray']);
	avgRowRange.setHorizontalAlignment('right');
	avgRowRange.setNumberFormat('###.#');
	
	let cellsToMerge = [
		`B${sortedData.length+1}:C${sortedData.length+1}`,
		'K3:L3', 'K4:L4', 'K5:L5', 'K6:L6', 'K7:L7',
	];

	let cellValues = [
		'none','Minutes in Wrap Up',
		'Minutes in Not Ready','Minutes in Call',
		'Minutes in Ready','Total Calls',
	];

	cellsToMerge.forEach( (val, ind) => {
		sheet.getRange(val).merge();
		if (cellValues[ind] != 'none') {
			sheet.getRange(val.split(":")[0]).setValue(cellValues[ind]);
			sheet.getRange(val.split(":")[0]).setHorizontalAlignment('center');
		} else {
			sheet.getRange(val.split(":")[0]).setHorizontalAlignment('right');
		}
	});

	let leftCol = [
		`=VLOOKUP($J$2,$B$2:$H$${sortedData.length+1},3,FALSE)`,
		`=VLOOKUP($J$2,$B$2:$H$${sortedData.length+1},4,FALSE)`,
		`=VLOOKUP($J$2,$B$2:$H$${sortedData.length+1},5,FALSE)`,
		`=VLOOKUP($J$2,$B$2:$H$${sortedData.length+1},6,FALSE)`,
		`=VLOOKUP($J$2,$B$2:$H$${sortedData.length+1},7,FALSE)`,
	];

	let rightCol = [
		`=VLOOKUP($M$2,$B$2:$H$${sortedData.length+1},3,FALSE)`,
		`=VLOOKUP($M$2,$B$2:$H$${sortedData.length+1},4,FALSE)`,
		`=VLOOKUP($M$2,$B$2:$H$${sortedData.length+1},5,FALSE)`,
		`=VLOOKUP($M$2,$B$2:$H$${sortedData.length+1},6,FALSE)`,
		`=VLOOKUP($M$2,$B$2:$H$${sortedData.length+1},7,FALSE)`,
	];

	let leftRange = addRange(sheetID, leftCol, 3, 10, sheet.getName(), 'col');
	let rightRange = addRange(sheetID, rightCol, 3, 13, sheet.getName(), 'col');
	rightRange.setHorizontalAlignment('left');

	let names = sheet.getRange(`B2:B${sortedData.length+1}`);
	let rule = SpreadsheetApp.newDataValidation().requireValueInRange(names).build();

	let leftCell = sheet.getRange('J2');
	let rightCell = sheet.getRange('M2');

	leftCell.setDataValidation(rule);
	rightCell.setDataValidation(rule);

	leftCell.setValue(sheet.getRange('B2').getValue());
	rightCell.setValue('AVERAGE');

	leftCell.setFontWeight('bold');
	rightCell.setFontWeight('bold');

	let comparisonCell = sheet.getRange('J28');
	comparisonCell.setValue('=J2&" vs "&M2');
	comparisonCell.setFontWeight('bold');

	let toWrite = [
		[
			'=J3-M3',
			'=IF(J29<0, "less than", "more than")',
			'=M2',
			'Minutes in Wrap Up',
		],
		[
			'=J4-M4',
			'=IF(J30<0, "less than", "more than")',
			'=M2',
			'Minutes in Not Ready',
		],
		[
			'=J5-M5',
			'=IF(J31<0, "less than", "more than")',
			'=M2',
			'Minutes in Call',
		],
		[
			'=J6-M6',
			'=IF(J32<0, "less than", "more than")',
			'=M2',
			'Minutes in Ready',
		],
		[
			'=J7-M7',
			'=IF(J32<0, "less than", "more than")',
			'=M2',
			'Total Calls',
		]
	];

	let lastRange = addRange(sheetID, toWrite, 29, 10, sheet.getName());

	return sheet;
}

function createToWriteArr(data) {
	// creates the array that will be written to the sheet
	let output = [];
	for (const name in data) {
		let row = [
			data[name]['login'],
			name,
			data[name]['email'],
			data[name]['inWrap'],
			data[name]['inNotReady'],
			data[name]['inCall'],
			data[name]['inReady'],
			data[name]['totalCalls'],
		  ];

		output.push(row);
	}

	output.pop(); // the last line is always empty, so remove it before sorting

	output = output.sort( (a,b) => {
		let nA = a[1].toLowerCase().charCodeAt();
		let nB = b[1].toLowerCase().charCodeAt();
		return nA - nB;
	});

	return output;
}

