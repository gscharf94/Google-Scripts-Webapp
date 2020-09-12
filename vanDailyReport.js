function getNamesVAN(fileID, fileName) {
	let newFile = convertToSheet(fileID);
	let names = readNames(newFile.id);
	return names;
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