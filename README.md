![workflow](https://github.com/do-/xlsxtream/actions/workflows/main.yml/badge.svg)
![Jest coverage](./badges/coverage-jest%20coverage.svg)

`xlsxtream` is a module for reading tabulated data from `.xlsx` (and, in general, [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML)) workbooks using node.js [streams API](https://nodejs.org/dist/latest/docs/api/stream.html). 

# Installation
```
npm install xlsxtream
```
# Usage
```js
const xlsx = require ('xlsxtream')

const wb = await xlsx.open ('/path/to/book.xlsx')

const ws = wb.sheetByName.Sheet1
// const ws = workbook.sheets [0] // WARNING: physical order

// just measuring 
const lastRow = await ws.getLastRow ()
if (lastRow.index > 10000) throw Error ('Size limit exceeded')

// extracting text, easy
const stringArrays = await ws.getObjectStream ()
for await (const [A, B, C] of stringArrays) {
  console.log (`A=${A}, B=${B}, C=${C}`)
  console.log (` (read ${worksheet.position} of ${worksheet.size} bytes)`)
}

// advanced processing
const rows = await ws.getRowStream ()
for await (const row of rows) {    // see https://github.com/do-/xlsxtream/wiki/Row
//console.log (`Logical #${row.index}, physical #${row.num}`)
  for (const cell of row.cells) {  // see https://github.com/do-/xlsxtream/wiki/Cell
    if (cell === null) {
      // empty cell, missing from XML
    }
    else {
      // process cell.index, cell.valueOf () etc.
    }
  }
}
```
# Rationale
Though [`exceljs`](https://github.com/exceljs/exceljs) is the standard _de facto_ for working with `.xlsx` files in the node.js ecosystem, it appears to be a bit too resource greedy in some applications. Specifically, when it takes to perform a simple data extraction (like converting to CSV) from a huge workbook (i. e. millions of rows). 

The original `exceljs` API is synchronous and presumes the complete data tree to be loaded in memory, which is totally inappropriate for such tasks. Since some point in time, the [Streaming I/O](https://github.com/exceljs/exceljs?tab=readme-ov-file#streaming-io) is implemented, but, as of writing this line, it still operates on a vast amount of data that may not be used at all.

The present module, `xlsxtream` takes the opposite approach: it uses the [streams API](https://nodejs.org/dist/latest/docs/api/stream.html) from the ground up, focuses on keeping the memory footprint as little as possible and letting the developer avoid every single operation not required by the task to perform. To this end, `xlsxtream`:
* reads individual files from ZIP archives with [`node-stream-zip`](https://www.npmjs.com/package/node-stream-zip);
* transforms them into filtered streams of XML nodes with [`xml-toolkit`](https://www.npmjs.com/package/xml-toolkit);
* provides [lazy iterators](https://github.com/do-/xlsxtream/wiki/Row) to access just the necessary cells' data.
