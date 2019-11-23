const xxx = require ('../index.js')
const fs  = require ('fs')

test ('scanSheetRows', async () => {
	
	let wb = await xxx.getWorkbook ((path) => fs.createReadStream ('./__data__/' + path))
	
	let last
	
	await xxx.scanSheetRows (wb, 'Лист1', row => last = row)
	
	expect (last [2].value).toBe ('2000-01-01')

})
