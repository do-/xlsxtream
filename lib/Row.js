const Cell = require ('./Cell.js')

class Row {

	constructor (node, sheet) {	

		this.node  = node

		this.sheet = sheet

	}

	get index () {

		return parseInt (this.node.attributes.get ('r'))

	}	

	toArrayOfStrings () {

		return this.node.children.map (c => new Cell (c, this).valueOf ())

	}
	
}

module.exports = Row