const {toYMD, toHMS, colIdx} = require ('./Utils.js')

class Cell {

	constructor (node, row) {	

		this.node = node

		this.row = row

	}

	get rawIndex () {

		return this.node.attributes.get ('r')

	}

	get index () {

		return colIdx (this.rawIndex.slice (0, - this.row.rawIndex.length))

	}

	valueOf () {

		const {node: {attributes, innerText}, row: {sheet: {book: {sharedStrings, numFmtId}}}} = this
		
		switch (attributes.get ('t')) {

			case 'n': 
			case 'inlineStr': 
				return innerText

			case 's': 
				return sharedStrings [parseInt (innerText)]

			default:

				const style = attributes.get ('s'); if (style) {
					
					switch (numFmtId [parseInt (style)]) {
		
						case '14':
						case '15':
							return toYMD (innerText)
		
						case '164':
							return toHMS (innerText)
		
					}
						
				}
		
				return innerText
	
		}

	}
	
}

module.exports = Cell