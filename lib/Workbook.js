const StreamZip = require ('node-stream-zip')
const {XMLParser, XMLReader, XMLNode} = require ('xml-toolkit')
const Worksheet = require ('./Worksheet.js')
const {toYMD, toHMS} = require ('./DateTimeUtils.js')

const parser = new XMLParser ()
const dump = XMLNode.toObject ()
const ZIP = Symbol ('zip')
const TODO = Symbol ('todo')

const TYPE_WORKSHEET      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const TYPE_SHARED_STRINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
const TYPE_STYLES         = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'

module.exports = class {

	constructor (path) {
	
		this.path = path
		this.numFmtId = []
		this.sharedStrings = []
		this.sheets = []
		this.sheetByName = {}
	
	}

	async load () {

		this [ZIP] = new StreamZip.async ({file: this.path})

		try {
		
			this [TODO] = [this.loadZipList ()]			
			await this.loadWorkbook ()
			await this.loadRelationships ()
			await Promise.all (this [TODO])

		}
		finally {

			await this [ZIP].close ()
			delete this [ZIP]

		}
			
	}
	
	async loadZipList () {

		this.zipList = await this [ZIP].entries ()

	}

	async parseXML (name) {

		const xml = await this [ZIP].entryData (name)

		return parser.process (xml.toString ())

	}

	async loadWorkbook () {

		this.sheetId2Name = new Map (); for (const {localName, children} of (await this.parseXML ('xl/workbook.xml')).children) if (localName === 'sheets') {

			for (const {attributes} of children)

				this.sheetId2Name.set (
					
					attributes.get ('id', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships'),

					attributes.get ('name')

				)

		}

	}

	async loadRelationships () {

		for (const {attributes} of (await this.parseXML ('xl/_rels/workbook.xml.rels')).children) {

			const target = attributes.get ('Target')

			switch (attributes.get ('Type')) {

				case TYPE_WORKSHEET:		
					const id = attributes.get ('Id')
					this.sheets.push (this.sheetByName [this.sheetId2Name.get (id)] = new Worksheet (this, id, target))
					break

				case TYPE_SHARED_STRINGS:
					this [TODO].push (this.loadSharedStrings (target))
					break
	
				case TYPE_STYLES:
					this [TODO].push (this.loadStyles (target))
					break

			}

		}

	}

	async loadStyles (path) {

		for (const i of (await this.parseXML ('xl/' + path)).children)

			for (const j of i.children) if (j.localName === 'xf' && j.attributes.has ('xfId'))

				this.numFmtId.push (j.attributes.get ('numFmtId'))

	}

	async loadSharedStrings (path) {

		const {sharedStrings} = this, zs = await this [ZIP].stream ('xl/' + path)
	
		return new Promise ((ok, fail) => {

			zs.on ('error', fail)

			new XMLReader ({filterElements : 't'})
				.process (zs)
				.on ('error', fail)
				.on ('end', ok)
				.on ('data', ({innerText}) => sharedStrings.push (innerText))

		})

	}
	
	getCellValue (c) {

		const d = dump (c), {v, t} = d
		
		switch (t) {
			case 's'         : return this.sharedStrings [parseInt (v)]
			case 'inlineStr' : return d.is.t
		}

		const {s} = d; if (s && !t) {
				
			switch (this.numFmtId [parseInt (s)]) {

				case '14':
				case '15':
					return toYMD (v)

				case '164':
					return toHMS (v)

			}
				
		}

		return v
	
	}

}