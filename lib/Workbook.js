const StreamZip = require ('node-stream-zip')
const {XMLParser, XMLReader, XMLNode} = require ('xml-toolkit')
const Worksheet = require ('./Worksheet.js')
const {toYMD, toHMS} = require ('./DateTimeUtils.js')

const dump = XMLNode.toObject ()
const ZIP = Symbol ('zip')

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
		
			const todo = [
				this.loadWorkbook (),
				this.loadZipList (),
			]
			
			await this.loadRelationships ()
			
			for (const {Type, Target} of this.relationships) switch (Type) {
			
				case TYPE_SHARED_STRINGS:
					todo.push (this.loadSharedStrings (Target))
					break

				case TYPE_STYLES:
					todo.push (this.loadStyles (Target))
					break
			
			}

			await Promise.all (todo)

		}
		finally {

			await this [ZIP].close ()

		}

		for (const r of this.relationships) if (r.Type === TYPE_WORKSHEET) {
		
			const s = new Worksheet (this, r.Id)
			
			this.sheets.push (s)
			
			for (const {name, id} of this.root.sheets.sheet) {
				
				if (r.Id !== id) continue
				
				this.sheetByName [name] = s
				
				break

			}
		
		}
		
	}
	
	async loadZipList () {

		this.zipList = await this [ZIP].entries ()

	}

	async parseXML (name, raw = false) {

		const xml = await this [ZIP].entryData (name)

		const root = new XMLParser ().process (xml.toString ())

		return raw ? root : dump (root)

	}

	async loadWorkbook () {

		const o = await this.parseXML ('xl/workbook.xml')

		if (!Array.isArray (o.sheets.sheet)) o.sheets.sheet = [o.sheets.sheet]

		this.root = o

	}

	async loadRelationships () {

		this.relationships = (await this.parseXML ('xl/_rels/workbook.xml.rels')).Relationship

	}

	async loadStyles (path) {

		for (const i of (await this.parseXML ('xl/' + path, true)).children)

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