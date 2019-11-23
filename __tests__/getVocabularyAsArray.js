const xxx = require ('../index.js')
const fs  = require ('fs')

test ('getVocabularyAsArray', async () => {
	
	expect (

		await xxx.getVocabularyAsArray (
		
			(path) => fs.createReadStream ('./__data__/' + path),

		)

	)
	.toEqual (

		expect.arrayContaining (['foo', 'bar', 'baz', 'no', 'label', 'date']),

    )
	
})
