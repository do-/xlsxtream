const xxx = require ('../index.js')
const fs  = require ('fs')

test ('scanStyles', async () => {

	let n = 0

	await xxx.scanStyles (
	
		(path) => fs.createReadStream ('./__data__/' + path),
		
		(node) => n += parseInt (node.attributes.numFmtId)
	
	)

	expect (n).toBe (14);

})
