const xxx = require ('../index.js')
const fs  = require ('fs')

test ('getStylesAsArray', async () => {
	
	expect (

		await xxx.getStylesAsArray (
		
			(path) => fs.createReadStream ('./__data__/' + path),

		)

	)
	.toEqual (

		[0, 14]

    )
	
})
