const {toYMD} = require ('../lib/Utils.js')

test ('basic', async () => {

	expect (toYMD (null)).toBeNull ()
	expect (toYMD ('42369')).toBe ('2015-12-31')
 	
})
