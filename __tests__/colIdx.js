const {colIdx} = require ('../lib/Utils.js')

test ('basic', async () => {

	expect (colIdx ('A4')).toBe (1)
	expect (colIdx ('T')).toBe (20)
	expect (colIdx ('Z11')).toBe (26)
	expect (colIdx ('AA4')).toBe (27)
	expect (colIdx ('AQ')).toBe (43)
	expect (colIdx ('BW')).toBe (75)
	expect (colIdx ('DF')).toBe (110)
	expect (colIdx ('AAA')).toBe (703)
 	
})
