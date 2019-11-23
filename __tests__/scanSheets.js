const xxx = require ('../index.js')
const fs  = require ('fs')

test ('scanSheets', async () => {

	let n = 0

	await xxx.scanSheets (
	
		(path) => fs.createReadStream ('./__data__/' + path),
		
		() => n ++
	
	)

	expect (n).toBe (3);

})
