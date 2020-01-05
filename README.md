# xlsxtream

This is a set of ES7 async functions for reading .xlsx files (Excel 2007) using the minimum resources possible.

The module operates on individual ZIP members in form of readable streams.

It doesen't rely on any specific unzipper. Its only exernal dependency is `sax`.

## Installation

```
npm install xlsxtream
```

## Basic usage

```
const xxx             = require ('xlsxtream')
const fs              = require ('fs')              // may not be needed...
// const zip_reader   = require ('node-stream-zip') // ...if you use this 
// const zip_reader   = require ('unzipper')        // ...or this

let streamProvider    = (path) => fs.createReadStream (root + path) 
// let streamProvider = await xxx.open (zip_reader, path) // if zip_reader is available

let wb = await xxx.getWorkbook (streamProvider) // not OO, just a plain js <<Object>>:
  // {
  //   streamProvider,
  //   saxOptions,
  //   sheets: {"Sheet One": "xl/worksheets/sheet1.xml", ...},
  //   styles: [0, 14, ...],         // from xl/styles.xml, to detect date/time cells
  //   stringResolver: s => voc [s], // from xl/sharedStrings.xml, to decode strings 
  //}

let [min RC, maxRC] = await xxx.getSheetDimensions (wb, "Sheet One") // [[1, 1], [1000, 3]]

await xxx.scanSheetRows (wb, "Sheet One", row => {
  // do something with 
  //  [
  //    {value: "1"}, 
  //    {value: "Text 1"}, 
  //  , ...]
})

```

## Explication

### streamProvider

Basically, .xlsx files are ZIP archives containing packs of XML documents.

`xlsxtream` operates on those XML directly. It's the developers's duty to provide it with the readable stream for any given local path.

For an .xlsx fully unzipped in a directory called `root`, the corresponding data source looks like

```
(path) => fs.createReadStream (root + path)
```

There is a plethora of ZIP related modules in the npm ecosystem. Anyone is free to unzip things the brand new way in each new project. The `xlsxtream` module depends on none of those libraries, though, has wrappers for two of them:
* https://www.npmjs.com/package/unzipper;
* https://www.npmjs.com/package/node-stream-zip;

```
let streamProvider = await xxx.open (zip_reader, path)
```

where `zip_reader` is the module reference (usually the result of `const ... = require ...`).

### Reading common workbook information

Having .xlsx content as a `streamProvider`, one can obtain the list if sheets and some other top workbook data with the call

```
let wb = await xxx.getWorkbook (streamProvider)

  // {
  //   streamProvider,
  //   saxOptions,
  //   sheets: {"Sheet One": "xl/worksheets/sheet1.xml", ...},
  //   styles: [0, 14, ...],         
  //   stringResolver: s => voc [s], 
  //}
  
```

The `sheets` component may be used for validation. All other are used when fetching sheet data.

For large workbooks, one may opt not to call this all-in-one method but to conctruct the object manually.

#### stringResolver

In .xlsx, string data from all sheets is stored externally: in `xl/sharedStrings.xml`.

So, to get text data from sheets cells, one need to first read this vocabulary file and to make its content available when parsing sheets.

The default, simplest way to do it is to buid an array of strings in memory:

```
 let voc = await xxx.getVocabularyAsArray (streamProvider)
 wb.stringResolver: index => voc [index]
```

Large files may require another option: to scan the vocabulary line by line to store it somewhere and then use the appropriate resolver:

```
 await xxx.scanVocabulary (streamProvider, (text, index) => {
   // store `text` at `index` (zero based int)
 })
 wb.stringResolver: index => // fetch by index
```

In an extreme case, `stringResolver` may be left undefined to join text data later (say, inside a database).

#### styles

Much like `xl/sharedStrings.xml`, each .xlsx file contains `xl/styles.xml` with style definitions.

It's crucial for decoding date/time data because such cells are marked only with style IDs.

The method for reading styles is

```
xxx.getStylesAsArray = async function (streamProvider, saxOptions = {}) {
  let a = []
  await xxx.scanStyles (streamProvider, node => 
    a.push (parseInt (node.attributes.numFmtId)
  ), saxOptions)
  return a
}
```

As one can see, the streaming mode is provided, but it's unlikely to be used directly because style lists are often limited in size:

```
wb.styles = await xxx.getStylesAsArray (streamProvider)
```

#### sheets

And, finally, consider the method to get the shhet list for a given workbook:

```
 wb.sheets = await xxx.getSheetsAsObject (streamProvider)
 if (not ('Sheet 1') in wb.sheets) throw 'Where is my Sheet 1?!'
```

Internally, it's too implemented the streaming way but one hardly ever need to override it.

### Measuring sheets

Having a workbook, one can get any of its sheets' dimensions:

```
let dim = await xxx.getSheetDimensions (wb, "Sheet One")                // with wb.sheets prefetched
let dim = await xxx.getSheetDimensions (wb, "xl/worksheets/sheet1.xml") // directly
// [
//   [1, 1],   // indices are 1-based
//   [1000, 3]
// ]
```

Technically, this information is read from an element at the very beginning of the XML document.

But the reading stream is closed right after it, so no resources are wasted to parse the rest.

### Reading sheet data

`xlsxtream` in intended for large .xslx workbooks. 

Such files most often contain sheets with long series of rows and fixed sets of columns.

That's why `xlsxtream` provides the row based API for reading sheet content:

```
await xxx.scanSheetRows (wb, "Sheet One", row => {
  // do something with 
  //  [
  //    {value: "1"}, 
  //    {value: "Text 1"}, 
  //  , ...]
})
```

The callback function receives an array containing an element for each cell in a row.

The `An` corresponds to th [0]th element, `Bn` to the [1]st one and so on. Regardless the global sheet dimension.

Empty cells may look like `null`s.

Non empty cells are copies of `sax` provided attribute sets for `<c>` nodes with attitional `value` components containing decoded content.

```
{
 r: 'A1',
 t: 's',             // it's a string
 v: '1',
 value: 'label'      // got from stringResolver
},
{
 r: 'B1',
 s: '14',            // say, this style is related to date/time
 v: '36526', 
 value: '2000-01-01' // `v` days since 1900-01-01
},
```

Normally, only the `value` is to be used in application code.

But, as mentioned before, one may opt to omit the `stringResolver` at parse time to join strings later or to use other low level optimizations.