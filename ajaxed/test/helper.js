var assertions = 1;
function assert(bool, msg) {
	var m = 'Success.';
	if (!bool) m = 'Failed: ' + ((msg) ? msg : 'no error help message given.');
	document.write(assertions + '. ' + m + '<br>');
	assertions++;
}