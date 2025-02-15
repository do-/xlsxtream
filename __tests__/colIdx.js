const {colIdx} = require ('../lib/Utils.js')

test ('basic', async () => {

	expect (colIdx ('A')).toBe (1)
	expect (colIdx ('T')).toBe (20)
	expect (colIdx ('Z')).toBe (26)
	expect (colIdx ('AA')).toBe (27)
	expect (colIdx ('AQ')).toBe (43)
	expect (colIdx ('BW')).toBe (75)
	expect (colIdx ('DF')).toBe (110)
	expect (colIdx ('AAA')).toBe (703)
 	
})
