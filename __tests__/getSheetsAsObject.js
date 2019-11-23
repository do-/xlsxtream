const xxx = require ('../index.js')
const fs  = require ('fs')

test ('getSheetsAsObject', async () => {

	expect (

		await xxx.getSheetsAsObject (
		
			(path) => fs.createReadStream ('./__data__/' + path),

		)

	)
	.toEqual (

		{
			'Лист1': 'xl/worksheets/sheet1.xml',
			'Лист2': 'xl/worksheets/sheet2.xml',
			'Лист3': 'xl/worksheets/sheet3.xml',
		},

    )	
	
})
