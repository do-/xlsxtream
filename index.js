const sax = require ('sax')

let xxx = module.exports = {}

const ordA = 'A'.charCodeAt (0) - 1

xxx.fromAZ = function (s) {

	let n = 0
	
	for (let i = 0; i < s.length; i ++) {	
		if (n) n *= 26		
		n += s.charCodeAt (i)
		n -= ordA
	}
	
	return n

}

xxx.toRC = function (s) {

	for (let i = 1; i < s.length; i ++) 
	
		if (s.charAt (i) <= '9') 
		
			return [parseInt (s.slice (i)), xxx.fromAZ (s.slice (0, i))]

}

xxx.scanVocabulary = async function (streamProvider, callBack, saxOptions = {}) {
	
	let reader = streamProvider ('xl/sharedStrings.xml')

	let i = 0, ss = sax.createStream (true, saxOptions)

	return new Promise ((ok, fail) => {

		let t

		ss.on ("opentag",  node => {if (node.name == 'si') t = ''})
		ss.on ("text",     text => t += text)
		ss.on ("closetag", name => {if (name == 'si') callBack (t, i ++)})

		reader.on   ('end', ok)
		reader.on   ('error', fail)
		
		reader.pipe (ss)

	})

}

xxx.getVocabularyAsArray = async function (streamProvider, saxOptions = {}) {

	let a = []
	
	await xxx.scanVocabulary (streamProvider, t => a.push (t), saxOptions)
	
	return a

}

xxx.scanStyles = async function (streamProvider, callBack, saxOptions = {}) {
	
	let reader = streamProvider ('xl/styles.xml')

	let i = 0, ss = sax.createStream (true, saxOptions)

	return new Promise ((ok, fail) => {

		ss.on ("opentag",  node => {if (node.name == 'xf' && ('xfId' in node.attributes)) {callBack (node)}})

		reader.on   ('end', ok)
		reader.on   ('error', fail)
		
		reader.pipe (ss)

	})

}

xxx.getStylesAsArray = async function (streamProvider, saxOptions = {}) {

	let a = []
	
	await xxx.scanStyles (streamProvider, node => a.push (parseInt (node.attributes.numFmtId)), saxOptions)
	
	return a

}

xxx.scanSheets = async function (streamProvider, callBack, saxOptions = {}) {

	let reader = streamProvider ('xl/workbook.xml')

	let i = 0, ss = sax.createStream (true, saxOptions)

	return new Promise ((ok, fail) => {

		ss.on ("opentag", node => {
			if (node.name == 'sheet') callBack (node.attributes)
		})

		reader.on   ('end', ok)
		reader.on   ('error', fail)		
		reader.pipe (ss)

	})

}

xxx.scanRels = async function (streamProvider, callBack, saxOptions = {}) {

	let reader = streamProvider ('xl/_rels/workbook.xml.rels')

	let i = 0, ss = sax.createStream (true, saxOptions)

	return new Promise ((ok, fail) => {

		ss.on ("opentag", node => {
			if (node.name == 'Relationship') callBack (node.attributes)
		})

		reader.on   ('end', ok)
		reader.on   ('error', fail)		
		reader.pipe (ss)

	})

}

xxx.getSheetsAsObject = async function (streamProvider, saxOptions = {}) {

	let o = {}
	
	let idx = {}
	
	await xxx.scanSheets (streamProvider, a => idx [a ['r:id']] = a.name, saxOptions)

	await xxx.scanRels (streamProvider, a => {if (a.Id in idx) o [idx [a.Id]] = 'xl/' + a.Target}, saxOptions)

	return o

}

xxx.getWorkbook = async function (streamProvider, saxOptions = {}) {

	let voc = await xxx.getVocabularyAsArray (streamProvider, saxOptions)

	return {
		streamProvider,
		saxOptions,
		sheets: await xxx.getSheetsAsObject (streamProvider, saxOptions),
		styles: await xxx.getStylesAsArray  (streamProvider, saxOptions),
		stringResolver: s => voc [s],
	}

}

xxx.getValue = function (workbook, cell) {

	let v = cell.v
	
	if (cell.t == 's' && workbook.stringResolver) v = workbook.stringResolver (parseInt (v))
	
	let s = cell.s; if (s && workbook.styles) switch (workbook.styles [parseInt (s)]) {
		case 14:
		case 15:
			let d = new Date ('1900-01-01')
			d.setDate (parseInt (v))
			v = d.toJSON ().slice (0, 10)
	}
		
	return v
	
}

xxx.getSheetDimensions = async function (workbook, name) {

	let path = (workbook.sheets || {}) [name] || name

	let reader = workbook.streamProvider (path)
	
	let ss = sax.createStream (true, workbook.saxOptions)
	
	return new Promise ((ok, fail) => {
		
		let dim
				
		ss.on ("opentag", node => {
			if (node.name == 'dimension') {
				let p = node.attributes.ref.split (':')
				if (p.length == 1) p [1] = p [0]
				dim = p.map (xxx.toRC)
				reader.destroy ()
				ok (dim)
			}
		})

		reader.on   ('end', ok)
		reader.on   ('error', (e) => dim ? ok (dim) : fail (e))

		reader.pipe (ss)
		
	})				

}

xxx.scanSheetRows = async function (workbook, name, callBack) {

	let path = (workbook.sheets || {}) [name] || name

	let reader = workbook.streamProvider (path)
	
	let ss = sax.createStream (true, workbook.saxOptions)
	
	return new Promise ((ok, fail) => {
		
		let row
		let t
		let cell
				
		ss.on ("opentag", node => {
			switch (node.name) {
				case 'row': return row = []
				case 'c'  : return cell = node.attributes
				case 'v'  : return t   = ''
			}
		})
		
		ss.on ("text",     text => t += text)
		
		ss.on ("closetag", name => {
			switch (name) {
				case 'row' : return callBack (row)
				case 'v'   : 
					cell.v = t
					let [r, c] = xxx.toRC (cell.r)
					cell.value = xxx.getValue (workbook, cell)
					row [c - 1] = cell
			}
		})
	
		reader.on   ('end', ok)
		reader.on   ('error', fail)
		
		reader.pipe (ss)
		
	})				

}
