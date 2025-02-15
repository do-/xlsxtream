class Row {

	constructor (node, sheet) {	

		this.node  = node

		this.sheet = sheet

	}

	get index () {

		return parseInt (this.node.attributes.get ('r'))

	}
	
}

module.exports = Row