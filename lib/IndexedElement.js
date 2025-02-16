class IndexedElement {

	constructor (node, parent) {	

		this.node = node

		this.parent = parent

	}

	get rawIndex () {

		return this.node.attributes.get ('r')

	}
	
}

module.exports = IndexedElement