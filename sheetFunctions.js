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

function addRange(sheetID, arr, row, col, sheetName, type="2d") {
	// adds arr to a range that anchored to row, col on top left
	// and then returns that range object
	let ss = SpreadsheetApp.openById(sheetID);
	let sheet = ss.getSheetByName(sheetName);
	
	if (type == "2d") {
		var sRow = row;
		var eRow = row+arr.length-1;
		var sCol = LETTERS[col];
		var eCol = LETTERS[col+arr[0].length-1];
	} else if (type == "col") {
		var sRow = row;
		var eRow = row+arr.length-1;
		var sCol = LETTERS[col];
		var eCol = LETTERS[col];

		var newArr = [];
		arr.forEach( (ele) => {
			newArr.push([ele]);
		});

		arr = newArr;
	} else {
		var sRow = row;
		var eRow = row;
		var sCol = LETTERS[col];
		var eCol = LETTERS[col+arr.length-1]
		arr = [arr];
	}
	let rangeText = `${sCol}${sRow}:${eCol}${eRow}`;
	let range = sheet.getRange(rangeText);
	range.setValues(arr);

	return range;
}

function createColorArr(width, height) {
    let arr = [];
    for(let i=0; i<height; i++) {
      let row = [];
      for(let j=0; j<width; j++) {
        if(i%2 == 0) {
          row.push('white');
        } else {
          row.push(COLORS['gray']);
        }
      }
      arr.push(row);
    }
    return arr;
  }