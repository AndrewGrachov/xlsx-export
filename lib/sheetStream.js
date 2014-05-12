var helpers = require("./helpers");
var util = require('util');
var stream = require('stream');
var sheetFront = '<?xml version="1.0" encoding="utf-8"?><x:worksheet xmlns:ss="urn:schemas-microsoft-com:office:spreadsheet" xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006" mc:Ignorable="x14ac" xmlns:x14ac="http://schemas.microsoft.com/office/spreadsheetml/2009/9/ac"><x:sheetPr><x:outlinePr summaryBelow="1" summaryRight="1" /></x:sheetPr><x:sheetViews><x:sheetView tabSelected="0" workbookViewId="0" /></x:sheetViews><x:sheetFormatPr defaultRowHeight="15" />';
var sheetBack = '<x:printOptions horizontalCentered="0" verticalCentered="0" headings="0" gridLines="0" /><x:pageMargins left="0.75" right="0.75" top="0.75" bottom="0.5" header="0.5" footer="0.75" /><x:pageSetup paperSize="1" scale="100" pageOrder="downThenOver" orientation="default" blackAndWhite="0" draft="0" cellComments="none" errors="displayed" /><x:headerFooter /><x:tableParts count="0" /></x:worksheet>';

getHeaderColumns = function (headers) {
    var cols = "";
	headers.forEach(function (header, keyIndex) {
		if(header.width) {
			cols += '<col customWidth = "1" width="' + header.width + '" max = "' + (keyIndex+1) +'" min="' + (keyIndex+1) +'"/>';
		}

	});
	if (cols !== "") {
		cols = '<x:cols>' + cols + '</x:cols>';
	}
	return cols;
}

getHeaderRow = function (headers) {
	var row = '<x:row r="1" spans="1:'+ headers.length + '">';
	var type = 'string';
	headers.forEach(function (header, index) {
		var cellRef = helpers.getColumnRef(index + 1, 0);
		row += helpers.typeHandlers[type](cellRef, header.caption);
	});
	row += '</x:row>';
	return row;
}

util.inherits(XMLStream, stream.Transform);

function XMLStream(options) {
	stream.Transform.call(this, {objectMode: true});
	this.map = options.map;

	this.headersLength = options.headers.length;
	this.rowIndex = 1;
	this.push(sheetFront);
	this.push(getHeaderColumns(options.headers));
	this.push('<x:sheetData>');
	this.push(getHeaderRow(options.headers));
}

XMLStream.prototype._transform = function (doc, encoding, callback) {
	var self = this;
	var row = '<x:row r="' + this.rowIndex + '" spans="1:' + this.headersLength + '">';
	Object.keys(this.map).forEach(function (mapKey, keyIndex) {
		var type = self.map[mapKey];
		var cellRef = helpers.getColumnRef(keyIndex + 1, self.rowIndex);
		row += helpers.typeHandlers[type](cellRef, doc[mapKey]);
	});

	row += '</x:row>';
	this.rowIndex += 1;
	this.push(row);
	callback();
}

XMLStream.prototype._flush = function (callback) {
	this.push('</x:sheetData>')
	this.push(sheetBack);
	this.rowIndex = 1;
	callback();
}

module.exports = XMLStream;