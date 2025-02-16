const {a1ToColIndex} = require ('../lib/Utils.js')

test ('basic', async () => {

	expect (a1ToColIndex ('A4')).toBe (1)
	expect (a1ToColIndex ('T')).toBe (20)
	expect (a1ToColIndex ('Z11')).toBe (26)
	expect (a1ToColIndex ('AA4')).toBe (27)
	expect (a1ToColIndex ('AQ')).toBe (43)
	expect (a1ToColIndex ('BW')).toBe (75)
	expect (a1ToColIndex ('DF')).toBe (110)
	expect (a1ToColIndex ('AAA')).toBe (703)
 	
})
