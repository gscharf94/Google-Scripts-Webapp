const COLORS = {
	'darkGray':'#7c7c7c',
	'gray':'#d7d7d7',
	'red':'#ff2525',
	'lightRed':'#f58080',
	'orange':'#ffc622',
	'yellow':'#fdf322',
	'green':'#08bd0e',
};

const LETTERS = [
	'NULL','A','B','C','D','E',
	'F','G','H','I','J','K',
	'L','M','N','O','P','Q',
	'R','S','T','U','V','W',
	'X','Y','Z'
];

function createSpreadsheet(originalName, parentFolder, type) {
	let parentID = parentFolder.getId();
	let today = new Date();
	let newName = `${originalName.slice(0,-4)} ${type}GENERATED ${today.getMonth() + 1}-${today.getDate()}`;

	let resource = {
		title: newName,
		mimeType: MimeType.GOOGLE_SHEETS,
		parents: [{ id: parentID }],
	};
	let fileJSON = Drive.Files.insert(resource);
	return fileJSON.id;
}

function parseCSV(rawText) {
	// turns CSV into 2d array
	let split = rawText.split("\n");
	let output = [];
	split.forEach( (row) => {
		let splitRow = row.split(",");
		output.push(splitRow);
	});
	return output;
}

function addRange(sheetID, arr, row, col, sheetName) {
	// adds arr to a range that anchored to row, col on top left
	// and then returns that range object
	let ss = SpreadsheetApp.openById(sheetID);
	let sheet = ss.getSheetByName(sheetName);
	let sRow = row;
	let eRow = row+arr.length-1;
	let sCol = LETTERS[col];
	let eCol = LETTERS[col+arr[0].length-1];
	let rangeText = `${sCol}${sRow}:${eCol}${eRow}`;
	let range = sheet.getRange(rangeText);
	range.setValues(arr);

	return range;
}