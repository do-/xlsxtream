const StreamZip             = require ('node-stream-zip')
const {Transform}           = require ('stream')
const {XMLReader}           = require ('xml-toolkit')
const  RowTransformer       = require ('./RowTransformer.js')
const  Row                  = require ('./Row.js')

module.exports = class {

	constructor (book, id, target) {
	
		this.book = book
		this.id   = id
		this.path = 'xl/' + target
		this.size = BigInt (book.zipList [this.path].size)
	
	}
	
	async getXMLStream () {
	
		this.position = 0n

		const 
			NOP = () => {}
			, zip = new StreamZip.async ({file: this.book.path})
			, zs = await zip.stream (this.path)

		zs.once ('close', () => zip.close ().then (NOP, NOP))
		
		const me = this, meter = new Transform ({
		
			transform (chunk, encoding, callback) {

				me.position += BigInt (chunk.length)
								
				callback (null, chunk)
			
			}
			
		})
		
		return zs.pipe (meter)

	}	

	async getXMLNodeStream () {
	
		const is = await this.getXMLStream ()

		return new XMLReader ({filterElements : 'row'}).process (is)

	}

	async getRowStream () {
	
		const is = await this.getXMLNodeStream ()

		let num = 0

		const me = this, xform = new Transform ({
			
			objectMode: true,
		
			transform (node, _, callback) {				

				const row = new Row (node, me)

				row.num = ++ num

				callback (null, row)
			
			}
			
		})

		return is.pipe (xform)

	}

	async getLastRow () {

		const rows = await this.getRowStream ()

		return new Promise ((ok, fail) => {

			let cur; rows
				.on ('error', fail)
				.on ('end', () => ok (cur))
				.on ('data', row => cur = row)

		})

	}
	
	async getObjectStream () {
	
		const is = await this.getXMLNodeStream ()
		
		const xform = new RowTransformer ({}, this)

		return is.pipe (xform)

	}

	toObject (xmlNode) {
	
		const {book} = this
	
		return xmlNode.children.map (c => book.getCellValue (c))
	
	}

}