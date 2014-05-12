var stream = require('stream');
var util = require('util');
var ZipPacker = require('../lib/xlsx-export');
var fs = require('fs');
var should = require('should');
var xlsx = require('xlsx');

util.inherits(Readable,stream.Readable);
function Readable() {
	stream.Readable.call(this, {objectMode: true});
}

Readable.prototype._read = function() {
	for(index; index < data.length; index++) {
		this.push(data[index]);
	}

	this.push(null);
};

var options = {
	map: {
		'date': 'date',
		'number': 'number',
		'name': 'string'
	},
	headers: [
				{width: 15, caption: "Awesome Date"},
				{width: 25, caption: 'Super awesome Number'},
				{width: 40, caption: 'Super Name'}
			],
	stream: new Readable()
};

var index = 0;

var data = [
{date: new Date(), number: 10, name: "Product1" },
{date: new Date(), number: 2, name: "Product2" },
{date: new Date(), number: 3, name: "Product3" },
{date: new Date(), number: 4, name: "Product4" },
{date: new Date(), number: 5, name: "Product5" },
{date: new Date(), number: 6, name: "Product6" }
];

describe('when creating xlsx spreadsheed', function () {
	before(function (done) {
		var zipPacker = new ZipPacker(options);
		var stream = zipPacker.pipe(fs.createWriteStream('./test_spread_sheet.xlsx'));
		stream.on('finish', function () {
			done();
		});
	});

	it('should create file', function () {
		fs.exists('./test_spread_sheet.xlsx', function (exists) {
			exists.should.equal(true);
		});
	});

	describe('when parsing created xlsx', function (done) {
		var sheet_name_list, workbook;
		before(function () {
			workbook = xlsx.readFile('./test_spread_sheet.xlsx', {cellHTML: true});
			sheet_name_list = workbook.SheetNames;
		});

		it('should read 1 sheet with name = "Sheet 1"', function () {
			expect(sheet_name_list).to.be.ok;
			expect(sheet_name_list.length).to.be.equal(1);
			expect(sheet_name_list[0]).to.be.equal('Sheet 1');
		});

		it('should parse something', function () {
			sheet_name_list.forEach(function(y) {
				var worksheet = workbook.Sheets[y];
				expect(worksheet.A2.t).to.equal('n');
				expect(worksheet.B2.w).to.equal('10');
				expect(worksheet.B2.t).to.equal('n');
			});
		});
	});
	after(function (done) {
		fs.unlink('./test_spread_sheet.xlsx', done);
	});
});