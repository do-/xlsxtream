# xlsxtream

This is a set of ES7 async functions for reading .xlsx files (Excel 2007; basically, zipped XML) using the minimum resources possible.

Instead of a ZIP archive, the module operates on its individual members in form of readable streams.

More specifically, xlsxtream uses a `streamProvider` (a function returning the stream given a local path) as its data source.

So, xlsxtream doesen't rely on any specific zip-related library. Its only exernal dependency is `sax`.

## Installation

```
npm install xlsxtream
```

## Basic usage

```
const xxx = require ('xlsxtream')
const fs  = require ('fs') // may not be needed

let streamProvider = (path) => fs.createReadStream (root + path) // or, you could use some unzipper here

let wb = await xxx.getWorkbook (streamProvider) // not OO, just a plain js <<Object>>:
/*
	{
		streamProvider,
		saxOptions,
		sheets: {"Sheet One": "xl/worksheets/sheet1.xml", ...},
		styles: [0, 14, ...],         // from xl/styles.xml, to detect date/time cells
		stringResolver: s => voc [s], // from xl/sharedStrings.xml, to decode strings 
	}
*/

let [min RC, maxRC] = await xxx.getSheetDimensions (wb, "Sheet One") // [[1, 1], [1000, 3]]

await xxx.scanSheetRows (wb, "Sheet One", row => {
  ... // do something with [{value: "1"}, null, {value: "Text 1"}, ...]
})

```
