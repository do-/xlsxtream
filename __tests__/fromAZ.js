const xxx = require ('../index.js')

test ('fromAZ', async () => {

	expect (xxx.fromAZ ('A')).toBe (1)
	expect (xxx.fromAZ ('Z')).toBe (26)
	expect (xxx.fromAZ ('AA')).toBe (27)
	expect (xxx.fromAZ ('AY')).toBe (51)
	expect (xxx.fromAZ ('ZZ')).toBe (702)
	expect (xxx.fromAZ ('AAC')).toBe (705)
	
})
