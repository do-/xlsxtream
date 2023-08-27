![workflow](https://github.com/do-/xlsxtream/actions/workflows/main.yml/badge.svg)
![Jest coverage](./badges/coverage-jest%20coverage.svg)

`xlsxtream` is a module for reading tabulated data from `.xlsx` (and, in general, [OOXML](https://en.wikipedia.org/wiki/Office_Open_XML)) workbooks using node.js [streams API](https://nodejs.org/dist/latest/docs/api/stream.html). 

Only common workbook data (including the shared strings list) are kept in memory; worksheets contents is read row by row using a chain of [transform streams](https://nodejs.org/dist/latest/docs/api/stream.html#class-streamtransform).

Basically, rows read are [Array](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/Array)s of [String](https://developer.mozilla.org/en-US/docs/Web/JavaScript/Reference/Global_Objects/String)s. Dates are [presented](https://github.com/do-/xlsxtream/wiki/DateTimeUtils) in `YYYY-MM-DD` format, time in `HH:MI:SS`. No effort is made to properly map scalar values to javaScript objects as the module is presumed to transform `.xlsx` data into [CSV](https://datatracker.ietf.org/doc/html/rfc4180) and other text formats for further bulk load into databases and similar operations.

Advanced users may prefer to alter the output format by accessing some lower level internals (see the reference below).

# Installation
```
npm install xlsxtream
```
# Usage
```js
const xlsx = require ('xlsxtream')

const workbook = xlsx.open ('/path/to/book.xlsx')

const worksheet = workbook.sheetByName.Sheet1
// const worksheet = workbook.sheets [0]
// worksheet.toObject = function (xmlNode) {return ...}

const rows = await ws.getObjectStream ()
/*
for await (const [A, B, C] of rows) {
  console.log (`A=${A}, B=${B}, C=${C}`)
  console.log (` (read ${worksheet.position} of ${worksheet.size} bytes)`)
}
*/
```

See the project's [wiki docs](https://github.com/do-/xlsxtream/wiki) for more information.
