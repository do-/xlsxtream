const Cell = require ('./Cell.js')

class Row {

	constructor (node, sheet) {	

		this.node  = node

		this.sheet = sheet

	}

	get rawIndex () {

		return this.node.attributes.get ('r')

	}	

	get index () {

		return parseInt (this.rawIndex)

	}

	* getCells () {

		let i = 1; for (const node of this.node.children) {

			const cell = new Cell (node, this), {index} = cell

			while (i < index) {

				yield null

				i ++

			}
			
			yield cell

			i ++

		}

	}

	toArrayOfStrings () {

		const a = [] 
		
		for (const c of this.getCells ()) a.push (c === null ? '' : c.valueOf ())

		return a

	}
	
}

module.exports = Row