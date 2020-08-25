function sortDictByTime(dict) {
	// sorts individual data by time
	// in descending order
	for (const name in dict) {
	  let data = dict[name];
	  let sortedData = data.sort(function(a,b) {
		return new Date('1970/01/01 ' + a[5]) - new Date('1970/01/01 ' + b[5]);
	  });
	  dict[name] = sortedData;
	}
	return dict;
}
