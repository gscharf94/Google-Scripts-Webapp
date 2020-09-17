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
				names[name[0]] = 0;
			}
		}
	}
	return [names, fileIds];
}

function createAverageHourlyStats(teamData, parentFolder) {
	

}

function getHourlyData(rawData) {
	let data = {};

	for (const time of rawData) {
		let hour = new Date(time).getHours();
		if (data[hour] == undefined) {
			data[hour] = 1;
		} else {
			data[hour]++;
		}
	}
	return data;
}

function createTimeSheet(rawData, parentFolder) {
	let ssid = createSpreadsheetNamed(parentFolder, 'Time Success');
	let ss = SpreadsheetApp.openById(ssid);
	let sheet = ss.getSheetByName('Sheet1');
	sheet.setName('Hourly');

	let hourlyData = getHourlyData(rawData);
	Logger.log(hourlyData);

	let rangeData = [];

	for (let i = 10; i < 22; i++) {
		if (hourlyData[i] != undefined) {
			rangeData.push( [`${i}:00`, String(hourlyData[i])] );
		}
	}

	let range = addRange(ssid, rangeData, 1, 1, 'Hourly');

	function createHourlyChart(sheet, range) {
		let chart = sheet.newChart();
		chart
			.addRange(range)
			.setChartType(Charts.ChartType.COLUMN)
			.setOption('width', 800)
			.setOption('height', 800)
			.setOption('legend', { position: "none" })
			.setOption('titlePosition', 'none')
			.setPosition(1,3,0,0);
		
		// let stackedChart = chart.asBarChart().setStacked();
		// let finishedChart = stackedChart.build();
		let finishedChart = chart.build();
		sheet.insertChart(finishedChart);
		return finishedChart;
	}

	createHourlyChart(sheet, range);
}

function createParentFolder(ids) {
	let parentFolder = DriveApp.getFileById(ids[0]).getParents().next();
	let newFolder = parentFolder.createFolder(`Time Success Report`);
	return newFolder;
}

function getTimeSuccess(ids) {
	
	let accumulated = [];
	for (const id of ids) {
		let ss = SpreadsheetApp.openById(id);
		let sheets = ss.getSheets().slice(2,);

		for (const sheet of sheets) {
			let times = getTimeSuccessFromSheet(sheet);
			Logger.log(`got ${times.length} times from ${sheet.getName()}`);
			for (const time of times) {
				accumulated.push(time);
			}
		}
	}

	return accumulated;
}

function getTimeSuccessFromSheet(sheet) {

	let length = sheet.getRange('A1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues().length;
	let timeCol = sheet.getRange(`C1:C${1+length}`).getValues();
	let canvCol = sheet.getRange(`H1:H${1+length}`).getValues();

	let success = [];

	for (let i = 0; i < timeCol.length; i++) {
		if (canvCol[i] == "X") {
			success.push(timeCol[i]);
		}
	}

	return success;
}


function startWeeklyReport(teams, idList, names) {
	

	let data = getTimeSuccess(idList);
	let newFolder = createParentFolder(idList);

	createTimeSheet(data, newFolder);
	// let teamData = {};

	// for (const team in teams) {
	// 	let pulledData  = getStats(teams, idList, teams[team], names);
	// 	teamData[team] = pulledData;
	// }

	// Logger.log(teamData);

	
}

function getStats(teams, idList, team, names) {
	let data = {};
	for (const name of team) {
		data[name] = {
			avgHourlyAttempts: 0,
			avgHourlyCanv: 0,
			minorTimeDiffs: 0,
		}

		for (const id of idList) {
			let stats = getStatsFromSheet(name, id);
			
			let sTime = new Date(stats[0]);
			let sTimeStr = `${String(sTime.getHours()).padStart(2, '0')}:${String(sTime.getMinutes()).padStart(2, '0')}`;
			let eTime = new Date(stats[1]);
			let eTimeStr = `${String(eTime.getHours()).padStart(2, '0')}:${String(eTime.getMinutes()).padStart(2, '0')}`;

			let hrsWorked = eTime - sTime;
			hrsWorked = (hrsWorked/60000)/60;

			let arr = [];

			for (const thing in stats[2]) {
				arr.push(thing);
			}
			data[name].minorTimeDiffs += arr.filter( (val) => val > 5 ).length;

			let overviewStats = getStatsFromOverview(name, id);
			data[name].avgHourlyAttempts += (overviewStats[0] / hrsWorked);
			data[name].avgHourlyCanv += (overviewStats[1] / hrsWorked);
		}
	}

	for (const name in data) {
		data[name].avgHourlyAttempts = (data[name].avgHourlyAttempts / names[name]);
		data[name].avgHourlyCanv = (data[name].avgHourlyCanv / names[name]);
		data[name].minorTimeDiffs = (data[name].minorTimeDiffs / names[name]);
	}

	return data;
}

function getStatsFromOverview(name, fileID) {
	let ss = SpreadsheetApp.openById(fileID);
	let sheet = ss.getSheets()[0];

	let nameCol = sheet.getRange("A1").getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let attemptsCol = sheet.getRange('C1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();
	let canvCol = sheet.getRange('D1').getDataRegion(SpreadsheetApp.Dimension.ROWS).getValues();

	let i;
	for (i = 0; i < nameCol.length; i++) {
		if (nameCol[i] == name) {
			return [ attemptsCol[i] , canvCol[i] ];
		}
	}

	return [0, 0];
}

function timeDiffs() {

}

function getStatsFromSheet(name, fileID) {

	let ss = SpreadsheetApp.openById(fileID);
	try {
		let sheet = ss.getSheetByName(name);
		let individualData = getIndividualTimeInfo(name, sheet);
		return individualData;
		Logger.log(`name: ${name} tDiffs: ${individualData[2]}`);

	} catch(err) {
		return [0,0,0,0];
	}
}