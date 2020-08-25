function callerDetailsStart(fileID, fileName) {
	let file = DriveApp.getFileById(fileID);
	let parentFolder = file.getParents().next();
	let blob = file.getBlob();
	let rawText = blob.getDataAsString('utf8');

	let parsedCSV = parseCSV(rawText);
	let data = createDict(parsedCSV);
	let sheetID = createSpreadsheet(fileName, parentFolder,'CD-');

	let overviewSheet = setupOverviewSheet(sheetID);
	
	let toWriteArr = createToWriteArr(data);
	
	let test = "";
	toWriteArr.forEach(
		(row) => {
			let tmp = "";
			row.forEach(
				(thing) => {
					tmp += `${thing} | `;
				}
				);
				test += `${tmp}<br>`;
			}
			);
			
	return test;
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

function setupOverviewSheet(sheetID) {
	// setups overview sheet
	// returns sheet obj
	let ss = SpreadsheetApp.openById(sheetID);
	let sheet = ss.getSheets()[0];
	sheet.setHiddenGridlines(true);

	let topRow = ['Login','Name','Email','Wrap Up','Not Ready','Call','Ready','Total'];
    
    let topRange = sheet.getRange('A1:H1');
    topRange.setValues([topRow]);
    topRange.setFontWeight('bold');
    topRange.setFontSize(11);
    topRange.setBackground(COLORS['darkGray']);
	topRange.setHorizontalAlignment('center');
	
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
	output = output.sort( (a,b) => {
		let nA = a[1].toLowerCase().charCodeAt();
		let nB = b[1].toLowerCase().charCodeAt();
		return nA - nB;
	});

	return output;
}

