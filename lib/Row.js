const IndexedElement = require ('./IndexedElement.js')
const Cell = require ('./Cell.js')

class Row extends IndexedElement {

	constructor (node, parent) {	

		super (node, parent)

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