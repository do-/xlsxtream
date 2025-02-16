const {toYMD, toHMS, a1ToColIndex} = require ('./Utils.js')
const IndexedElement = require ('./IndexedElement.js')

class Cell extends IndexedElement {

	get index () {

		return a1ToColIndex (this.rawIndex)

	}

	valueOf () {

		const {node: {attributes, innerText}, parent: {parent: {book: {sharedStrings, numFmtId}}}} = this
		
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