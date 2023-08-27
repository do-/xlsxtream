const zip = require ('unzippo')
const {XMLParser, XMLReader, XMLNode} = require ('xml-toolkit')
const Worksheet = require ('./Worksheet.js')
const {toYMD, toHMS} = require ('./DateTimeUtils.js')

const dump = XMLNode.toObject ()

const TYPE_WORKSHEET      = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet'
const TYPE_SHARED_STRINGS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings'
const TYPE_STYLES         = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles'

const MS_IN_DAY = 24 * 60 * 60 * 1000

module.exports = class {

	constructor (path) {
	
		this.path = path
		this.numFmtId = []
		this.sharedStrings = []
		this.name2path = new Map ()
		this.sheets = []
		this.sheetByName = {}
	
	}

	async load () {
		
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

		this.zipList = await zip.list (this.path)

	}

	async loadWorkbook () {

		const xml = await zip.read (this.path, 'xl/workbook.xml')

		const o = XMLNode.toObject () (new XMLParser ().process (xml))
		
		if (!Array.isArray (o.sheets.sheet)) o.sheets.sheet = [o.sheets.sheet]

		this.root = o

	}

	async loadRelationships () {

		const xml = await zip.read (this.path, 'xl/_rels/workbook.xml.rels')
				
		this.relationships = XMLNode.toObject () (new XMLParser ().process (xml)).Relationship

	}

	async loadStyles (path) {

		const xml = await zip.read (this.path, 'xl/' + path)
				
		for (const i of (new XMLParser ().process (xml)).children)

			for (const j of i.children) if (j.localName === 'xf' && j.attributes.has ('xfId'))

				this.numFmtId.push (j.attributes.get ('numFmtId'))

	}

	async loadSharedStrings (path) {
	
		const {sharedStrings} = this, zs = await zip.open (this.path, 'xl/' + path)
	
		return new Promise ((ok, fail) => {

			zs.on ('error', fail)

			const os = new XMLReader ({filterElements : 't'}).process (zs)

			os.on ('error', fail)
			os.on ('end', ok)

			os.on ('data', ({children}) => {

				let t = ''; for (const i of children) t += i.text
			
				sharedStrings.push (t)

			})

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