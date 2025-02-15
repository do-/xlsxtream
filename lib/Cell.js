class Cell {

	constructor (node, row) {	

		this.node = node

		this.row = row

	}

	valueOf () {

		return this.row.sheet.book.getCellValue (this.node)

	}
	
}

module.exports = Cell