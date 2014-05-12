xlsx-export
===========

very basic node js streaming json -> xlsx converter

**test**

```
npm test
```

**Usage**

```

var XlsxExport = require('xlsx-export');
var collection = mongo.collection('someCollection');

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
	stream: collection.stream()
};

var packer = new XlsxExport(options);
var stream = packer.pipe(fs.createWriteStream('./test_spread_sheet.xlsx'));

//todo: enjoy!

```
