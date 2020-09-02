function startThruText(fileID, fileName) {
	let file = DriveApp.getFileById(fileID);
	let parentFolder = file.getParents().next();
	let blob = file.getBlob();
	let rawText = blob.getDataAsString('utf8');
	let parsedCSV = parseCSV(rawText);

	let ssID = createSpreadsheet(fileName, parentFolder, "TT-");
	
	let sortedData = categorizeData(parsedCSV);
	
	let callsRange = writeCallsSheet(ssID, sortedData[2]);
	let inc = mergeIncoming(sortedData[1]);
	let incs = splitIncoming(inc);
	let incRanges = writeIncomingSheets(ssID, incs);
	formatIncomingSheets(ssID, incRanges, incs);


	writeOverviewSheet(ssID, sortedData[0], incs, sortedData[2]);
	
	// let out = sortedData[0];
	// let calls = sortedData[2];
	
}

function writeOverviewSheet(ssID, outgoing, incs, calls) {
	let ss = SpreadsheetApp.openById(ssID);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Overview');
	sheet.setHiddenGridlines(true);

	let personCounter = {};
	outgoing.forEach( (row) => {
		let name = `${row[4]} ${row[5]}`;
		personCounter[name] = 0;
	});
	
	outgoing.forEach( (row) => {
		let name = `${row[4]} ${row[5]}`;
		personCounter[name]++;
	});

	let toWrite = [[
		'Name',
		'Outgoing',
	]];

	for (const person in personCounter) {
		let row = [
			person,
			personCounter[person],
		];
		toWrite.push(row);
	}

	let outgoingRange = addRange(ssID, toWrite, 1, 1, 'Overview');

	let colors = createColorArr(2, toWrite.length-1);
	colors.unshift([
		COLORS['darkGray'],
		COLORS['darkGray'],
	]);

	outgoingRange.setBackgrounds(colors);

	let outgoingTopRange = sheet.getRange('A1:B1');
	outgoingTopRange.setFontWeight('bold');
	outgoingTopRange.setHorizontalAlignment('center');

	let nextRow = toWrite.length+2;

	function getCallBacks() {
		let count = 0;
		calls.forEach( (row) => {
			if(row[2] == "call") {
				count++;
			}
		});
		return count;
	}

	let comparisonData = [
		[
			'Total Outgoing',
			`${outgoing.length}`,
			'',
		],
		[
			'Opt Outs',
			`${(incs[0].length-1)}`,
			`${(incs[0].length-1)/outgoing.length}`,
		],
		[
			'Wrong Numbers',
			`${(incs[1].length-1)}`,
			`${(incs[1].length-1)/outgoing.length}`,
		],
		[
			'Call Backs',
			`${getCallBacks()}`,
			`${getCallBacks()/outgoing.length}`,
		],
	];

	let comparisonRange = addRange(ssID, comparisonData, nextRow, 1, 'Overview');
	comparisonRange.setFontWeight('bold');

	let leftRange = sheet.getRange(`A${nextRow}:A${nextRow+3}`);
	leftRange.setFontWeight('bold');
	leftRange.setHorizontalAlignment('right');

	let percentRange = sheet.getRange(`C${nextRow+1}:C${nextRow+3}`);
	percentRange.setNumberFormat("0.0%");

	function createFirstChart() {
		let range = outgoingRange;
		let chart = sheet.newChart();
		chart
			.addRange(range)
			.setChartType(Charts.ChartType.COLUMN)
			.setOption('width', 600)
			.setOption('height', 300)
			// .setOption('legend', { position: "none" })
			.setOption('titlePosition', 'none')
			.setPosition(1,3,0,0);
		
		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart;
	}

	let firstChart = createFirstChart();
	let secondChart = createSecondChart();

	function createSecondChart() {
		let range = sheet.getRange(`A${nextRow}:B${nextRow+3}`);
		let chart = sheet.newChart();
		chart
			.addRange(range)
			.setChartType(Charts.ChartType.PIE)
			.setOption('width', 450)
			.setOption('height', 350)
			.setOption('legend', { position: "right" })
			.setOption('titlePosition', 'none')
			.setPosition(nextRow+4,1,0,0);
		
		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart;
	}

}

function writeCallsSheet(ssID, calls) {
	let ss = SpreadsheetApp.openById(ssID);
	ss.insertSheet('CALLS');
	let sheet = ss.getSheetByName('CALLS');

	let toWrite = [[
		'Contact Name',
		'Contact ID',
		'Message',
		'Type',
		'MessageID',
		'Convo ID',
		'Convo Phone',
		'Contact Phone'
	]];
	for (let i=1; i<calls.length-1; i++) {
		let fpn = calls[i][15];
		let fPhoneNum = `(${fpn.slice(2,5)}) ${fpn.slice(5,8)}-${fpn.slice(8,12)}`;
		let cpn = calls[i][11];
		let cPhoneNum = `(${cpn.slice(2,5)}) ${cpn.slice(5,8)}-${cpn.slice(8,12)}`;

		let row = [
			`${calls[i][13]} ${calls[i][14]}`,
			calls[i][12],
			calls[i][3],
			calls[i][2],
			calls[i][1],
			calls[i][10],
			cPhoneNum,
			fPhoneNum,
		];

		toWrite.push(row);
	}

	let mainRange = addRange(ssID, toWrite, 1, 1, 'CALLS');

	let colors = createColorArr(8, toWrite.length-1);
	colors.unshift([
		COLORS['darkGray'], COLORS['darkGray'],
		COLORS['darkGray'], COLORS['darkGray'],
		COLORS['darkGray'], COLORS['darkGray'],
		COLORS['darkGray'], COLORS['darkGray'],
	]);

	mainRange.setBackgrounds(colors);

	let topRowRange = sheet.getRange('A1:H1');
	topRowRange.setFontWeight('bold');
	topRowRange.setHorizontalAlignment('center');

	sheet.setHiddenGridlines(true);

	sheet.setColumnWidth(1, 140);
	sheet.setColumnWidth(2, 140);
	sheet.setColumnWidth(3, 230);
	sheet.setColumnWidth(4, 60);
	sheet.setColumnWidth(5, 140);
	sheet.setColumnWidth(6, 140);
	sheet.setColumnWidth(7, 93);
	sheet.setColumnWidth(8, 93);




	return mainRange;
}

function writeIncomingSheets(ssID, incs) {
	let ss = SpreadsheetApp.openById(ssID);
	
	let topRow = ["Contact Name", "Message", "Contact ID", "Convo Phone","Contact Phone"]
	incs[0].unshift(topRow)
	incs[1].unshift(topRow)
	incs[2].unshift(topRow)
	
	ss.insertSheet('OPT OUT');
	let sheet = ss.getSheetByName('OPT OUT');
	let stopRange = addRange(ssID, incs[0], 1, 1, 'OPT OUT');

	ss.insertSheet('WRONG NUMBER');
	sheet = ss.getSheetByName('WRONG NUMBER');
	let wrongNumRange = addRange(ssID, incs[1], 1, 1, 'WRONG NUMBER');

	ss.insertSheet('TRUMP SUPPORTER');
	sheet = ss.getSheetByName('TRUMP SUPPORTER');
	let trumpRange = addRange(ssID, incs[2], 1, 1, 'TRUMP SUPPORTER');

	return [stopRange, wrongNumRange, trumpRange];
}

function formatIncomingSheets(ssID, ranges, incs) {
	let ss = SpreadsheetApp.openById(ssID);
	let sheetNames = ['OPT OUT','WRONG NUMBER','TRUMP SUPPORTER'];
	ranges.forEach( (range, ind) =>{
		let data = incs[ind];
		let sheet = ss.getSheetByName(sheetNames[ind]);
		let colors = createColorArr(5, data.length-1);
		colors.unshift([
			COLORS['darkGray'], COLORS['darkGray'],
			COLORS['darkGray'], COLORS['darkGray'],
			COLORS['darkGray'],
		]);
		range.setBackgrounds(colors);
		range.setWrap(true);
		sheet.setHiddenGridlines(true);

		let topRowRange = sheet.getRange('A1:E1');
		topRowRange.setFontWeight('bold');
		topRowRange.setHorizontalAlignment('center');

		sheet.setColumnWidth(1, 140);
		sheet.setColumnWidth(2, 275);
		sheet.setColumnWidth(3, 140);
		sheet.setColumnWidth(4, 95);
		sheet.setColumnWidth(5, 95);

		for (let i=2; i<data.length+1; i++) {
			sheet.setRowHeight(i, 21);
		}

	});
}

function categorizeData(arr) {
	let [out, inc, calls] = getOutInMessages(arr);

	return [out, inc, calls];
}


function getOutInMessages(arr) {
	let outgoing = [];
	let incoming = [];
	let calls = [];
	arr.forEach( (row) => {
		if (row[2] == "outgoing") {
			let rowCopy = [ ...row ]
			outgoing.push(rowCopy);
		} else if (row[2] == "incoming") {
			let rowCopy = [ ...row ];
			incoming.push(rowCopy);
		} else {
			let rowCopy = [ ...row ];
			calls.push(rowCopy);
		}
	});
	return [outgoing, incoming, calls];
}


function mergeIncoming(inc) {
	let incObj = {};

	inc.forEach( (row) => {
		let name = `${row[13]} ${row[14]}`;
		incObj[name] = {
			convoPhone: '',
			phoneNum: '',
			contactID: '',
			messages: [],
			convoID: '',
		};
	});

	inc.forEach( (row) => {
		let message = row[3];
		let convoID = row[10];
		let phoneNum = row[11];
		let contactID = row[12];
		let contactName = `${row[13]} ${row[14]}`;
		let contactPhone = row[15];

		incObj[contactName]['phoneNum'] = phoneNum;
		incObj[contactName]['messages'].push(message);
		incObj[contactName]['contactID'] = contactID;
		incObj[contactName]['convoID'] = convoID;
		incObj[contactName]['convoPhone'] = contactPhone;
	});
	return incObj;
}

function splitIncoming(incs) {
	let stops = [];
	let wrongNum = [];
	let trump = [];
	for (const person in incs) {
		let msgs = incs[person]['messages'];
		let mergedMsg = "";
		for (let i=0; i<msgs.length; i++) {
			mergedMsg += msgs[i];
		}
		mergedMsg = mergedMsg.toLowerCase();
		let fpn = incs[person]['phoneNum'];
		let fPhoneNum = `(${fpn.slice(2,5)}) ${fpn.slice(5,8)}-${fpn.slice(8,12)}`;
		let cpn = incs[person]['convoPhone'];
		let cPhoneNum = `(${cpn.slice(2,5)}) ${cpn.slice(5,8)}-${cpn.slice(8,12)}`;

		let toAppend = [
			person,
			mergedMsg,
			incs[person]["convoID"],
			fPhoneNum,
			cPhoneNum,
		];

		if (mergedMsg.indexOf('stop') != -1) {
			stops.push(toAppend);
		}

		if (
			mergedMsg.indexOf('wrong') != -1 ||
			mergedMsg.indexOf('equivocado') != -1
			) {
			wrongNum.push(toAppend);
		}

		if (mergedMsg.indexOf('trump') != -1) {
			trump.push(toAppend);
		}
	}
	return [stops, wrongNum, trump]
}

