const {toHMS} = require ('../lib/DateTimeUtils.js')

test ('basic', async () => {

	expect (toHMS (null)).toBeNull ()
	expect (toHMS ('0.99930555555555556')).toBe ('23:59:00')
 	
})
