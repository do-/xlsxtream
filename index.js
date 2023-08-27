const Workbook = require ('./lib/Workbook.js')

module.exports = {

	open: async path => {
	
		const book = new Workbook (path)
		
		await book.load ()
		
		return book
	
	}

}