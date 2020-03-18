const Saxophone = require ('saxophone')

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

xxx.scan = async function (streamProvider, path, handler) {

	let reader = await streamProvider (path)

	let ss = new Saxophone ()		
	
	ss.__stream = reader

	return new Promise ((ok, fail) => {
	
		for (let event in handler) ss.on (event, handler [event])

		reader.on   ('error', fail)
		reader.on   ('end', ok)
		reader.on   ('close', ok)
		
		reader.pipe (ss)

	})	

}

xxx.scanVocabulary = async function (streamProvider, callBack) {

	let t = '', i = 0

	return xxx.scan (streamProvider, 'xl/sharedStrings.xml', {

		tagopen:  node => {if (node.name == 'si') t = ''},

		text:     text => t += Saxophone.parseEntities (text.contents),

		tagclose: node => {if (node.name == 'si') callBack (t, i ++)},		

	})
	
}

xxx.getVocabularyAsArray = async function (streamProvider) {

	let a = []
	
	await xxx.scanVocabulary (streamProvider, t => a.push (t))
	
	return a

}

xxx.scanStyles = async function (streamProvider, callBack) {

	return xxx.scan (streamProvider, 'xl/styles.xml', {

		tagopen:  node => {if (node.name == 'xf' && ('xfId' in Saxophone.parseAttrs (node.attrs))) {callBack (node)}},

	})

}

xxx.getStylesAsArray = async function (streamProvider) {

	let a = []
	
	await xxx.scanStyles (streamProvider, node => a.push (parseInt (Saxophone.parseAttrs (node.attrs).numFmtId)))
	
	return a

}

xxx.scanSheets = async function (streamProvider, callBack) {

	return xxx.scan (streamProvider, 'xl/workbook.xml', {

		tagopen:  node => {if (node.name == 'sheet') callBack (Saxophone.parseAttrs (node.attrs))},

	})

}

xxx.scanRels = async function (streamProvider, callBack) {

	return xxx.scan (streamProvider, 'xl/_rels/workbook.xml.rels', {

		tagopen:  node => {if (node.name == 'Relationship') callBack (Saxophone.parseAttrs (node.attrs))},

	})

}

xxx.getSheetsAsObject = async function (streamProvider) {

	let o = {}
	
	let idx = {}
	
	await xxx.scanSheets (streamProvider, a => idx [a ['r:id']] = a.name)

	await xxx.scanRels (streamProvider, a => {if (a.Id in idx) o [idx [a.Id]] = 'xl/' + a.Target})

	return o

}

xxx.getWorkbook = async function (streamProvider) {

	if (typeof streamProvider === 'object') {
	
		let o = streamProvider

		let clazz = o.constructor.name
	
		switch (clazz) {
		
			case 'StreamZip':			
				streamProvider = p => new Promise ((ok, fail) => o.stream (p, (x, s) => x ? fail (x) : ok (s)))			
				break

			case 'unzipper':			
				streamProvider = p => {
					for (let file of o.dir.files) if (file.path == p) return file.stream ()
				}			
				break
			
			default:
				throw 'Unknown unzipper: ' + clazz
		
		}
	
	}

	let voc = await xxx.getVocabularyAsArray (streamProvider)

	return {
		streamProvider,
		sheets: await xxx.getSheetsAsObject (streamProvider),
		styles: await xxx.getStylesAsArray  (streamProvider),
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
			d.setMinutes (d.getTimezoneOffset ())
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

		tagopen: function (node) {

			if (node.name == 'dimension') {
				let p = Saxophone.parseAttrs (node.attrs).ref.split (':')
				if (p.length == 1) p [1] = p [0]
				callback (p.map (xxx.toRC))
				this.__stream.destroy ()
			}
			
		},

	})

}

xxx.scanSheetRows = async function (workbook, name, callBack) {

	let row
	let t
	let cell
	
	return xxx.scan (workbook.streamProvider, (workbook.sheets || {}) [name] || name, {

		tagopen: node => {switch (node.name) {
		
			case 'row': return row = []
			
			case 'c'  : return cell = Saxophone.parseAttrs (node.attrs)
			
			case 'v'  : return t   = ''
		
		}},

		text:     text => t += Saxophone.parseEntities (text.contents),

		tagclose: async node => {switch (node.name) {
		
			case 'row' : return callBack (row)
			
			case 'v'   : 
				cell.v = t
				let [r, c] = xxx.toRC (cell.r)
				cell.value = xxx.getValue (workbook, cell)
				row [c - 1] = cell
				
		}},

	})
	
}

xxx.openStreamZip = async function (zip, file) {

	let zf = (await new Promise ((ok, fail) => {
	
		let z = new zip ({
			file,
			storeEntries: true,
			skipEntryNameValidation: true,
		})
		.on ('error', fail)
		.on ('ready', () => ok (z))
		
	})).on ('error', x => console.log (x))
	
	return zf
	
}

xxx.open = async function (zip, file) {

	let clazz = zip.name
	
	if (!clazz) {
		if (zip.Open && zip.Open.s3) clazz = 'unzipper'
	}
	
	switch (clazz) {
		case 'StreamZip': return xxx.openStreamZip (zip, file)
		case 'unzipper': return {
			constructor: {name: 'unzipper'},
			dir: await zip.Open.file (file),
		}
		default: throw 'Unknown unzipper class: ' + clazz
	}

}