function doGet() {
	return HtmlService
		.createTemplateFromFile('page')
		.evaluate();
}

function getInitialFolders() {
	let output = [];
	let rootFolder = DriveApp.getRootFolder();
	let folders = rootFolder.getFolders();
	while (folders.hasNext()) {
		output.push(folders.next());
	}
	// return formatListAsHTML(output);
	return output;
}


function formatListAsHTML(list) {
	let output = '<p>';
	list.forEach(
		(val) => {
			output += `${val}<br>`;
		}
	);
	output += '</p>';
	return output;
}

function dropdownSelection(folderName) {
	let folder = DriveApp.getFoldersByName(folderName).next();
	let folders = folder.getFolders();
	let files = folder.getFiles();

	let html = "";
	while (folders.hasNext()) {
		let folder = folders.next();
		html += `<li class="folderLink" onclick="selectFolder('${folder}')">${folder}</li>`;
	}
	while (files.hasNext()) {
		html += `<li>${files.next()}</li>`;
	}

	return html;
}

function include(filename) {
	return HtmlService.createHtmlOutputFromFile(filename)
		.getContent();
}