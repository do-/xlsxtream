const {toYMD, toHMS} = require ('./DateTimeUtils.js')

class Cell {

	constructor (node, row) {	

		this.node = node

		this.row = row

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