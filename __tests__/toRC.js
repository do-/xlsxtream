const xxx = require ('../index.js')

test ('toRC', async () => {

	expect (xxx.toRC ('A1')).toEqual ([1, 1])
	expect (xxx.toRC ('A10')).toEqual ([10, 1])
	expect (xxx.toRC ('AA1')).toEqual ([1, 27])

})
