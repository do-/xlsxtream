const Cell = require ('./Cell.js')

class Row {

	constructor (node, sheet) {	

		this.node  = node

		this.sheet = sheet

		Object.defineProperty (this, 'cells', {configurable: true, enumerable: true,

			get: function * () {

				let i = 1; for (const node of this.node.children) {

					const cell = new Cell (node, this), {index} = cell

					while (i < index) { yield null; i ++ }

					yield cell; i ++

				}

			}

		})

	}

	get rawIndex () {

		return this.node.attributes.get ('r')

	}	

	get index () {

		return parseInt (this.rawIndex)

	}

	toArrayOfStrings () {

		const a = [] 
		
		for (const c of this.cells) a.push (c === null ? '' : c.valueOf ())

		return a

	}
	
}

module.exports = Row