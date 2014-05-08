module.exports = {
	getColumnRef: function(col, cellIndex) {
	  cellIndex= cellIndex + 1;
	  if (col <= 0)
		throw "col must be more than 0";
	  var array = new Array();
	  while (col > 0)
	  {
		var remainder = col % 26;
		col /= 26;
		col = Math.floor(col);
		if(remainder === 0)
		{
			remainder = 26;
			col--;
		}
		array.push(64 + remainder);
	  }
	  var letter = String.fromCharCode.apply(null, array.reverse());
	  return letter + cellIndex.toString();
	},

	typeHandlers: {
		string : function(cellRef, value, styleIndex) {
			styleIndex = styleIndex || 0;
			if (value===null || value === undefined)
				return "";
			if (typeof value ==='string') {
				value = value.replace(/&/g, "&amp;").replace(/'/g, "&apos;").replace(/>/g, "&gt;").replace(/</g, "&lt;");
			}
			return '<x:c r="'+cellRef+'" s="'+ styleIndex + '" t="inlineStr"><x:is><x:t>'+ value + '</x:t></x:is></x:c>';
		},
		number: function(cellRef, value, styleIndex) {
			styleIndex = styleIndex || 0;
			if (value === null || value === undefined) {
				return "";
			}
			return '<x:c r="' + cellRef + '" s="'+ styleIndex +'" t="n"><x:v>' + value + '</x:v></x:c>';
		},
		date: function(cellRef, value, styleIndex) {
			styleIndex = styleIndex || 1;
			if (value === null || value === undefined) {
				return "";
			}
			//trasform js date to excel day counter from 30/11/1899
			var DAY = 24 * 60 * 60 * 1000;
			var originDate = new Date(Date.UTC(1899, 11, 30));
			value = (value - originDate) / DAY;

			return '<x:c r="' + cellRef + '" s="' + styleIndex + '" t="n"><x:v>' + value + '</x:v></x:c>';
		},
		boolean: function(cellRef, value, styleIndex) {
			styleIndex = styleIndex || 0;
				if (value === null || value === undefined) {
					return "";
				}
				if (value) {
				  value = 1;
				} else {
				  value = 0;
				}

				return '<x:c r="'+ cellRef + '" s="'+ styleIndex + '" t="b"><x:v>'+value+'</x:v></x:c>';
			}
		}
	}