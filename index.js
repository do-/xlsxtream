const sax = require ('sax')

let xxx = module.exports = {}

xxx.fromAZ = function (s) {

	const ordA = 'A'.charCodeAt (0) - 1
	
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

xxx.scan = async function (streamProvider, path, handler, saxOptions = {}) {

	let reader = streamProvider (path)

	let ss = sax.createStream (true, saxOptions)
	
	ss.__stream = reader

	return new Promise ((ok, fail) => {
	
		for (let event in handler) ss.on (event, handler [event])

		reader.on   ('close', ok)
		reader.on   ('error', fail)
		
		reader.pipe (ss)

	})	

}

xxx.scanVocabulary = async function (streamProvider, callBack, saxOptions = {}) {

	let t = '', i = 0

	return xxx.scan (streamProvider, 'xl/sharedStrings.xml', {

		opentag:  node => {if (node.name == 'si') t = ''},

		text:     text => t += text,

		closetag: name => {if (name == 'si') callBack (t, i ++)},		

	}, saxOptions)
	
}

xxx.getVocabularyAsArray = async function (streamProvider, saxOptions = {}) {

	let a = []
	
	await xxx.scanVocabulary (streamProvider, t => a.push (t), saxOptions)
	
	return a

}

xxx.scanStyles = async function (streamProvider, callBack, saxOptions = {}) {

	return xxx.scan (streamProvider, 'xl/styles.xml', {

		opentag:  node => {if (node.name == 'xf' && ('xfId' in node.attributes)) {callBack (node)}},

	}, saxOptions)

}

xxx.getStylesAsArray = async function (streamProvider, saxOptions = {}) {

	let a = []
	
	await xxx.scanStyles (streamProvider, node => a.push (parseInt (node.attributes.numFmtId)), saxOptions)
	
	return a

}

xxx.scanSheets = async function (streamProvider, callBack, saxOptions = {}) {

	return xxx.scan (streamProvider, 'xl/workbook.xml', {

		opentag:  node => {if (node.name == 'sheet') callBack (node.attributes)},

	}, saxOptions)

}

xxx.scanRels = async function (streamProvider, callBack, saxOptions = {}) {

	return xxx.scan (streamProvider, 'xl/_rels/workbook.xml.rels', {

		opentag:  node => {if (node.name == 'Relationship') callBack (node.attributes)},

	}, saxOptions)

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
	let dim
	await xxx.scanSheetDimensions (workbook, name, (v => dim = v))
	return dim
}

xxx.scanSheetDimensions = async function (workbook, name, callback) {

	return xxx.scan (workbook.streamProvider, (workbook.sheets || {}) [name] || name, {

		opentag: function (node) {

			if (node.name == 'dimension') {
				let p = node.attributes.ref.split (':')
				if (p.length == 1) p [1] = p [0]
				callback (p.map (xxx.toRC))
				this.__stream.destroy ()
			}
			
		},

	}, workbook.saxOptions)

}

xxx.scanSheetRows = async function (workbook, name, callBack) {

	let row
	let t
	let cell
	
	return xxx.scan (workbook.streamProvider, (workbook.sheets || {}) [name] || name, {

		opentag: node => {switch (node.name) {
		
			case 'row': return row = []
			
			case 'c'  : return cell = node.attributes
			
			case 'v'  : return t   = ''
		
		}},

		text:     text => t += text,

		closetag: async name => {switch (name) {
		
			case 'row' : return await callBack (row)
			
			case 'v'   : 
				cell.v = t
				let [r, c] = xxx.toRC (cell.r)
				cell.value = xxx.getValue (workbook, cell)
				row [c - 1] = cell
				
		}},

	}, workbook.saxOptions)
	
}
