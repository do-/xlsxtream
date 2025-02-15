const StreamZip             = require ('node-stream-zip')
const {Transform}           = require ('stream')
const {XMLReader}           = require ('xml-toolkit')
const  RowTransformer       = require ('./RowTransformer.js')

module.exports = class {

	constructor (book, id) {
	
		this.book = book
		this.id   = id
		this.path = 'xl/' + book.relationships.find (i => i.Id === id).Target
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