class Row {

	constructor (node, sheet) {	

		this.node  = node

		this.sheet = sheet

	}

	get index () {

		return parseInt (this.node.attributes.get ('r'))

	}

	toArrayOfStrings () {

		const {node, sheet: {book}} = this

		return node.children.map (c => book.getCellValue (c))

	}
	
}

module.exports = Row