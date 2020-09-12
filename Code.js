function doGet() {
	return HtmlService
		.createTemplateFromFile('page')
		.evaluate();
}

function getRootInfo() {
	let rootFolder = DriveApp.getRootFolder();
	let folderIterator = rootFolder.getFolders();
	let folders = [];
	while (folderIterator.hasNext()) {
		folders.push(folderIterator.next());
	}
	let fileIterator = rootFolder.getFiles();
	let files = [];
	while (fileIterator.hasNext()) {
		files.push(fileIterator.next());
	}

	let html = "";
		
	folders.forEach( (folder) => {
		html += `<li class="folderLink" onclick="folderForwardAction('${folder}')">${folder}</li>`;
	});

	files.forEach( (file) => {
		if(`${file}`.slice(-3) == 'csv') {
			html += `<li class="fileLinkCSV" onclick="fileSelectionAction('${file}','My Drive','${file.getId()}')">${file}</li>`;
		} else if (`${file}`.slice(-4) == 'xlsx') {
			html += `<li class="fileLinkXLSX" onclick="fileSelectionAction('${file}','My Drive','${file.getId()}')">${file}</li>`;
		}
	});

	return { 'folders':folders, 'files':files, 'html':html };
}


function dropdownSelection(folderName) {
	let folder = DriveApp.getFoldersByName(folderName).next();
	let folders = folder.getFolders();
	let files = folder.getFiles();
	let html = "";
	while (folders.hasNext()) {
		let folder = folders.next();
		html += `<li class="folderLink" onclick="folderForwardAction('${folder}')">${folder}</li>`;
	}
	while (files.hasNext()) {
		let file = files.next();
		if(`${file}`.slice(-3) == 'csv') {
			html += `<li class="fileLinkCSV" onclick="fileSelectionAction('${file}','My Drive','${file.getId()}')">${file}</li>`;
		} else if (`${file}`.slice(-4) == 'xlsx') {
			html += `<li class="fileLinkXLSX" onclick="fileSelectionAction('${file}','My Drive','${file.getId()}')">${file}</li>`;
		}
	}
	return {'name':folderName, 'html':html};
}

function getParentInfo(folderName) {
	let parentFolder = DriveApp.getFoldersByName(folderName).next().getParents().next();

	let folders = parentFolder.getFolders();
	let files = parentFolder.getFiles();

	let html = "";
	while (folders.hasNext()) {
		let folder = folders.next();
		html += `<li class="folderLink" onclick="folderForwardAction('${folder}')">${folder}</li>`;
	}
	while (files.hasNext()) {
		let file = files.next();
		if(`${file}`.slice(-3) == 'csv') {
			html += `<li class="fileLinkCSV" onclick="fileSelectionAction('${file}','My Drive','${file.getId()}')">${file}</li>`;
		} else if (`${file}`.slice(-4) == 'xlsx') {
			html += `<li class="fileLinkXLSX" onclick="fileSelectionAction('${file}','My Drive','${file.getId()}')">${file}</li>`;
		}
	}
	return {'name':folderName, 'html':html};
}

function getUserInfo() {
	let data = Drive.About.get();
	let name = data['user']['displayName'];
	let imageURL = data['user']['picture']['url'];

	let output = {
		'email':Session.getActiveUser().getEmail(),
		'name': name,
		'imageURL':imageURL,
	};
	return output;

}

function include(filename) {
	return HtmlService
		.createHtmlOutputFromFile(filename)
		.getContent();
}