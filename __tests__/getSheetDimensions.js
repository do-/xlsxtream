const xxx = require ('../index.js')
const fs  = require ('fs')

test ('getSheetDimensions', async () => {
	
	let wb = await xxx.getWorkbook ((path) => fs.createReadStream ('./__data__/' + path))

	expect (
		await xxx.getSheetDimensions (wb, 'Лист1')
	)
	.toEqual (
		[
			[1, 1],
			[4, 3],
		]
    )

	expect (
		await xxx.getSheetDimensions (wb, 'Лист2')
	)
	.toEqual (
		[
			[1, 1],
			[1, 1],
		]
    )
    
})
