const {Transform} = require ('stream')

module.exports = class extends Transform {

	constructor (options, sheet) {

		options.objectMode = true

		super (options)

		this.sheet = sheet

	}
	
	_transform (chunk, encoding, callback) {
		
		try {
		
			const row = this.sheet.toObject (chunk)

			callback (null, row)

		}
		catch (error) {

			callback (error)

		}

	}

}