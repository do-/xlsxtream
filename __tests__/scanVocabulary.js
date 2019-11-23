const xxx = require ('../index.js')
const fs  = require ('fs')

test ('scanVocabulary', async () => {

	let n = ''

	await xxx.scanVocabulary (
	
		(path) => fs.createReadStream ('./__data__/' + path),
		
		(t, i) => n += i
	
	)

	expect (n).toBe ('012345');

})
