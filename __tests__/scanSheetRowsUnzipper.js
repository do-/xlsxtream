const xxx = require ('../index.js')
const fs  = require ('fs')
const uz  = require ('unzipper')

test ('scanSheetRows', async () => {

	let zip = await xxx.open (uz, './__data__/test.xlsx')
		
	let wb = await xxx.getWorkbook (zip)
	
	let last
	
	await xxx.scanSheetRows (wb, 'Лист1', row => last = row)
	
	expect (last [2].value).toBe ('2000-01-01')

})
