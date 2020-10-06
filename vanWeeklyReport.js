function getWeeklyNames(fileID) {
	let parentFolder = DriveApp.getFileById(fileID).getParents().next();
	let names = {};
	let fileIds = [];
	let parentFolderFiles = parentFolder.getFiles();
	while (parentFolderFiles.hasNext()) {
		let file = parentFolderFiles.next();
		let newFile = convertToSheet(file.getId());
		fileIds.push(newFile.getId());
		let fileNames = readNames(newFile.getId());
		for (const name of fileNames) {
			if (names[name[0]] != undefined) {
				names[name[0]]++;
			} else {
				names[name[0]] = 1;
			}
		}
	}
	return [names, fileIds];
}


function startWeeklyReport(teams, idList, names) {

	let timeSuccessData = {};
	let teamStats = {};
	for (const team in teams) {
		teamStats[team] = {};
	}
	let individualStats = {};
	for (const name in names) {
		individualStats[name] = {};
	}

	let statsAggregate = {};

	for (const ssid of idList) {
		let [dailyStats, dailyCanvTimes, dailyIndividualData, dateStr] = getDataFromSpreadsheet(ssid);
		timeSuccessData[dateStr] = dailyCanvTimes;
		statsAggregate[dateStr] = dailyStats;
		for (const name in dailyIndividualData) {
			if (individualStats[name] == undefined) {
				individualStats[name] = {};
			} else {
				individualStats[name][dateStr] = dailyIndividualData[name];
			}
		}
	}

	// createTimeSuccessCharts(timeSuccessData, idList[0]);
	createTeamStatsWeeklyReport(statsAggregate, idList[0], individualStats, teams);
	// historicalWeeklyIndividualReport(individualStats, idList[0], statsAggregate);
	return 'complete';
}

function createTeamStatsWeeklyReport(statsAggregate, childFileId, individualStats, teams) {
	// need average hourly stats per team
	// need total stacked bar graph

	let parentFolder = DriveApp.getFileById(childFileId).getParents().next();
	let ssid = createSpreadsheetNamed(parentFolder, 'Weekly Team Stats');

	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Team Response Breakdown');

	let teamData = {};
	let template = {
		total: 0,
		canv: 0,
		notHome: 0,
		refused: 0,
		disconnected: 0,
		moved: 0,
	};

	for (const team in teams) {
		Logger.log(`creating team: ${team}`);
		teamData[team] = {
			...template
		};
		for (const person of teams[team]) {
			for (const day in statsAggregate) {
				for (const person2 in statsAggregate[day]) {
					if (person2 == person) {
						teamData[team].total += statsAggregate[day][person2].totalCalls;
						teamData[team].canv += statsAggregate[day][person2].canv;
						teamData[team].notHome += statsAggregate[day][person2].notHome;
						teamData[team].refused += statsAggregate[day][person2].refused;
						teamData[team].disconnected += statsAggregate[day][person2].disconnected;
						teamData[team].moved += statsAggregate[day][person2].moved;
					}
				}
			}
		}
	}

	Logger.log(teamData);
}

function historicalWeeklyIndividualReport(individualStats, childFileId, statsAggregate) {
	// new one should be interesting..

	let parentFolder = DriveApp.getFileById(childFileId).getParents().next();
	let ssid = createSpreadsheetNamed(parentFolder, 'Weekly Individual Stats');

	let ss = SpreadsheetApp.openById(ssid);

}

function createTimeSuccessCharts(timeSuccessData, childFileId) {
	let parentFolder = DriveApp.getFileById(childFileId).getParents().next();
	let ssid = createSpreadsheetNamed(parentFolder, `Weekly Time Success`);

	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Time Success')

	let hourlySplit = {};
	for (const day in timeSuccessData) {
		hourlySplit[day] = {};
		for (const time of timeSuccessData[day]) {
			let hour = time.getHours();
			if (hourlySplit[day][hour] == undefined) {
				hourlySplit[day][hour] = 1;
			} else {
				hourlySplit[day][hour]++;
			}
		}
	}

	Logger.log(hourlySplit);

	let dataToWrite = formatDataArr(hourlySplit);
	let heatmapRange = addRange(ssid, dataToWrite, 4, 2, 'Time Success');

	let max = 0;
	for (const day in hourlySplit) {
		for (const hour in hourlySplit[day]) {
			if (hourlySplit[day][hour] > max) {
				max = hourlySplit[day][hour];
			}
		}
	}

	let dataArr = [];
	for (let i = 0; i < dataToWrite.length - 1; i++) {
		let row = [];
		for (let j = 1; j < dataToWrite[i].length; j++) {
			row.push(dataToWrite[i][j]);
		}
		dataArr.push(row);
	}

	let percentages = [];

	for (const row of dataArr) {
		let pRow = [];
		for (const val of row) {
			pRow.push(val / max);
		}
		percentages.push(pRow);
	}

	let backgrounds = [];

	for (const row of percentages) {
		let bRow = [];
		for (const val of row) {
			if (val > .8) {
				bRow.push('#08183A');
			} else if (val > .6) {
				bRow.push('#152852');
			} else if (val > .4) {
				bRow.push('#4B3D60');
			} else if (val > .2) {
				bRow.push('#FD5E53');
			} else {
				bRow.push('#FC9C54');
			}
		}
		backgrounds.push(bRow);
	}

	let numHours = dataToWrite[0].length - 1;
	let numDays = dataToWrite.length - 1;

	Logger.log(`C4:${LETTERS[2+numHours]}${3+numDays}`);
	let dataRange = sheet.getRange(`C4:${LETTERS[2+numHours]}${3+numDays}`);
	dataRange
		.setBackgrounds(backgrounds)
		.setFontSize(14)
		.setFontColor('white')
		.setFontWeight('bold')
		.setHorizontalAlignment('center')
		.setVerticalAlignment('middle');

	for (let i = 4; i < 4 + numDays; i++) {
		sheet.setRowHeight(i, 41);
	}

	for (let i = 3; i < 3 + numHours; i++) {
		sheet.setColumnWidth(i, 53);
	}

	sheet.setHiddenGridlines(true);

	sheet.getRange('C2')
		.setValue('Weekly Time Success Heatmap')
		.setFontWeight('bold')
		.setFontSize(12);

	sheet.getRange('B1:B15')
		.setFontWeight('bold')
		.setFontSize(12)
		.setHorizontalAlignment('center');

	sheet.getRange(`C${4+numDays}:${LETTERS[2+numHours]}${4+numDays}`)
		.setFontWeight('bold')
		.setFontSize(12)
		.setHorizontalAlignment('center');

	createTimeSuccessHourlyChart(sheet, hourlySplit, ssid);
}

function createTimeSuccessHourlyChart(sheet, hourlySplit, ssid) {
	let totals = {};
	let hoursArr = [];
	for (const day in hourlySplit) {
		for (const hour in hourlySplit[day]) {
			let val = hourlySplit[day][hour];
			if (totals[hour] == undefined) {
				totals[hour] = val;
			} else {
				totals[hour] += val;
			}
			if (hoursArr.indexOf(hour) == -1) {
				hoursArr.push(hour);
			}
		}
	}

	hoursArr.sort((a, b) => {
		return Number(a) - Number(b);
	});

	let toWrite = [];
	for (const hour of hoursArr) {
		toWrite.push([
			hour,
			totals[hour],
		]);
	}

	let range = addRange(ssid, toWrite, 15, 1, 'Time Success');

	let chart = sheet.newChart();
	chart
		.addRange(range)
		.setChartType(Charts.ChartType.COLUMN)
		.setOption('width', 600)
		.setOption('height', 600)
		.setOption('legend', {
			position: 'none'
		})
		.setOption('titlePosition', 'none')
		.setOption('useFirstColumnAsDomain', 'true')
		.setOption('series.0.dataLabel', 'value')
		.setPosition(15, 1, 0, 0);

	let finishedChart = chart.build();
	sheet.insertChart(finishedChart);

	sheet.getRange('A12')
		.setValue('Time Success Hourly Totals')
		.setFontWeight('bold')
		.setFontSize(14);
}

function formatDataArr(data) {
	let daysArr = [];
	for (const day in data) {
		daysArr.push(day);
	}

	daysArr.sort((a, b) => {
		let aMonth = Number(a.split("-")[0]);
		let aDay = Number(a.split("-")[1]);
		let bMonth = Number(b.split("-")[0]);
		let bDay = Number(b.split("-")[1]);

		let aSum = (aMonth * 30) + aDay;
		let bSum = (bMonth * 30) + bDay;
		return aSum - bSum;
	});

	let hours = [];

	for (const day in data) {
		for (const hour in data[day]) {
			if (hours.indexOf(hour) == -1) {
				hours.push(hour);
			}
		}
	}

	hours.sort((a, b) => {
		return Number(a) - Number(b);
	});

	let output = [];

	for (const day of daysArr) {
		let row = [day];
		for (const hour of hours) {
			let val = data[day][hour];
			if (val == undefined) {
				row.push(0);
			} else {
				row.push(val);
			}
		}
		output.push(row);
	}

	let lastRow = [''];
	for (const hour of hours) {
		lastRow.push(hour);
	}

	output.push(lastRow);
	return output;
}

function getDataFromSpreadsheet(ssid) {
	let ss = SpreadsheetApp.openById(ssid);
	let sheets = ss.getSheets();

	let date = ss.getName().slice(0, 6);
	let dateStr = `${date.slice(0,2)}-${date.slice(2,4)}-${date.slice(4,6)}`;
	Logger.log(`starting day: ${dateStr}`);

	let overviewSheet = sheets.shift();
	let bigSheet = sheets.shift();

	let totalCanvTimes = [];
	let indivData = {};

	for (const sheet of sheets) {
		let [hoursWorked, canvTimes, timeDiffs] = getIndividualData(sheet);
		for (canvTime of canvTimes) {
			totalCanvTimes.push(canvTime);
		}
		indivData[sheet.getSheetName()] = {
			hoursWorked: hoursWorked,
			timeDiffs: timeDiffs,
		};
	}

	let dailyInfo = getOverviewSheetDataWeekly(overviewSheet);

	// add hrly calls & canv to indivData
	for (const name in indivData) {
		let totalCalls = dailyInfo[name].totalCalls;
		let totalCanv = dailyInfo[name].canv;

		indivData[name].hourlyCalls = totalCalls / indivData[name].hoursWorked;
		indivData[name].hourlyCanv = totalCanv / indivData[name].hoursWorked;
	}

	return [dailyInfo, totalCanvTimes, indivData, dateStr];
}

function getOverviewSheetDataWeekly(sheet) {
	let lastRow = sheet.getLastRow();

	let nameCol = sheet.getRange(`A1:A${lastRow}`).getValues();
	let callsCol = sheet.getRange(`C1:C${lastRow}`).getValues();
	let canvCol = sheet.getRange(`D1:D${lastRow}`).getValues();
	let refCol = sheet.getRange(`F1:F${lastRow}`).getValues();
	let notHomeCol = sheet.getRange(`J1:J${lastRow}`).getValues();
	let disCol = sheet.getRange(`K1:K${lastRow}`).getValues();
	let movedCol = sheet.getRange(`L1:L${lastRow}`).getValues();

	let peopleInfo = {};

	for (let i = 0; i < nameCol.length; i++) {
		let name = nameCol[i][0];
		peopleInfo[name] = {
			totalCalls: callsCol[i][0],
			canv: canvCol[i][0],
			refused: refCol[i][0],
			notHome: notHomeCol[i][0],
			disconnected: disCol[i][0],
			moved: movedCol[i][0],
		};
	}

	return peopleInfo;
}

function getIndividualData(sheet) {
	let lastRow = sheet.getLastRow();

	let startTime = sheet.getRange('C2').getValue();
	let endTime = sheet.getRange(`C${lastRow}`).getValue();
	let hoursWorked = (endTime - startTime) / 60000 / 60;

	let timeCol = sheet.getRange(`C2:C${lastRow}`).getValues();
	let canvCol = sheet.getRange(`H2:H${lastRow}`).getValues();

	let canvTimes = [];

	for (let i = 0; i < timeCol.length; i++) {
		let time = timeCol[i][0];
		time.setHours(time.getHours() - 3);
		let canv = canvCol[i][0];
		if (canv == 'X') {
			canvTimes.push(time);
		}
	}

	let allTimeDiffs = getTimeDiffs(timeCol);
	let minorTimeDiffCount = 0;
	for (const diff of allTimeDiffs) {
		if (diff > 6) {
			minorTimeDiffCount++;
		}
	}

	return [hoursWorked, canvTimes, minorTimeDiffCount];
}